using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace NovelProofreader
{
    public class Issue
    {
        public int Paragraph;
        public int Start;
        public int Length;
        public string Type;
        public string Original;
        public string Suggestion;
        public string Note;
        public int Priority;

        public string Location
        {
            get { return (Paragraph + 1) + "段" + (Start + 1) + "字"; }
        }
    }

    public class ScanOptions
    {
        public bool Punctuation = true;
        public bool Advanced = true;
        public bool Style = true;
        public bool Names = true;
        public bool Online = false;
    }

    public class ScanResult
    {
        public List<Issue> Issues = new List<Issue>();
        public int ParagraphCount;
        public int CharacterCount;
        public bool OnlineUsed;
        public string OnlineError;
        public int OnlineIssueCount;
    }

    internal class PhraseRule
    {
        public string Wrong;
        public string Right;
        public string Type;
        public string Note;
    }

    internal class RegexRule
    {
        public Regex Pattern;
        public string Replacement;
        public string Type;
        public string Note;
        public int Priority;
        public bool PunctuationOnly;
        public bool AdvancedOnly;
        public bool StyleOnly;
    }

    public class Scanner
    {
        public const string Typo = "错别字";
        public const string Lex = "词法";
        public const string Style = "用词建议";
        public const string Redundant = "冗余";
        public const string Punctuation = "标点";
        public const string Name = "名称";

        public const string Author = "小说校对助手";
        public const string ShadingColor = "FFF2A8";

        private static readonly List<PhraseRule> PhraseRules = BuildPhraseRules();
        private static readonly List<RegexRule> ContextRules = BuildContextRules();

        public ScanResult Scan(List<string> paragraphs, ScanOptions options)
        {
            if (options == null) options = new ScanOptions();
            ScanResult result = new ScanResult();
            result.ParagraphCount = paragraphs.Count;

            List<Issue> candidates = new List<Issue>();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                string text = paragraphs[i] ?? string.Empty;
                result.CharacterCount += text.Length;
                if (text.Length == 0) continue;
                ScanParagraph(text, i, options, candidates);
            }

            if (options.Names)
            {
                candidates.AddRange(ScanNameConsistency(paragraphs));
            }

            result.Issues = Resolve(candidates);
            return result;
        }

        public static List<Issue> MergeIssues(List<Issue> first, List<Issue> second)
        {
            List<Issue> all = new List<Issue>();
            if (first != null) all.AddRange(first);
            if (second != null) all.AddRange(second);
            return Resolve(all);
        }

        private static void ScanParagraph(string text, int para, ScanOptions options, List<Issue> output)
        {
            // Exact wrong phrase / word dictionary.
            foreach (PhraseRule rule in PhraseRules)
            {
                if (rule.Type == Style && !options.Style) continue;
                int index = 0;
                while (index < text.Length)
                {
                    int hit = text.IndexOf(rule.Wrong, index, StringComparison.Ordinal);
                    if (hit < 0) break;
                    Issue issue = new Issue();
                    issue.Paragraph = para;
                    issue.Start = hit;
                    issue.Length = rule.Wrong.Length;
                    issue.Type = rule.Type;
                    issue.Original = rule.Wrong;
                    issue.Suggestion = rule.Right;
                    issue.Note = string.IsNullOrEmpty(rule.Note) ? DefaultNote(rule.Type, rule.Right) : rule.Note;
                    issue.Priority = 100;
                    output.Add(issue);
                    index = hit + rule.Wrong.Length;
                }
            }

            // Context-sensitive regex rules.
            foreach (RegexRule rule in ContextRules)
            {
                if (rule.PunctuationOnly && !options.Punctuation) continue;
                if (rule.AdvancedOnly && !options.Advanced) continue;
                if (rule.StyleOnly && !options.Style) continue;
                foreach (Match match in rule.Pattern.Matches(text))
                {
                    if (!match.Success || match.Length == 0) continue;
                    string suggestion = match.Result(rule.Replacement);
                    if (suggestion == match.Value) continue;
                    Issue issue = new Issue();
                    issue.Paragraph = para;
                    issue.Start = match.Index;
                    issue.Length = match.Length;
                    issue.Type = rule.Type;
                    issue.Original = match.Value;
                    issue.Suggestion = suggestion;
                    issue.Note = string.IsNullOrEmpty(rule.Note) ? DefaultNote(rule.Type, suggestion) : rule.Note;
                    issue.Priority = rule.Priority;
                    output.Add(issue);
                }
            }

            // Adjacent duplicate function characters.
            foreach (Match match in DuplicateCharRegex.Matches(text))
            {
                if (match.Length < 2) continue;
                Issue issue = new Issue();
                issue.Paragraph = para;
                issue.Start = match.Index;
                issue.Length = match.Length;
                issue.Type = Typo;
                issue.Original = match.Value;
                issue.Suggestion = match.Value.Substring(0, 1);
                issue.Note = "相邻重复字，疑似笔误；请确认是否为对话中的口吃。";
                issue.Priority = 85;
                output.Add(issue);
            }

            if (options.Punctuation)
            {
                AddParagraphEndPunctuation(text, para, output);
            }
        }

        private static void AddParagraphEndPunctuation(string text, int para, List<Issue> output)
        {
            string trimmed = text.TrimEnd();
            if (trimmed.Length < 20) return;
            char last = trimmed[trimmed.Length - 1];
            if (!IsHan(last)) return;

            // Headings such as "第一章 归来" normally have no full stop.
            if (trimmed.Length <= 28 && Regex.IsMatch(trimmed, @"^第[一二三四五六七八九十百千0-9]+[章节卷回部篇]")) return;

            int index = text.LastIndexOf(last);
            Issue issue = new Issue();
            issue.Paragraph = para;
            issue.Start = index;
            issue.Length = 1;
            issue.Type = Punctuation;
            issue.Original = last.ToString();
            issue.Suggestion = last + "。";
            issue.Note = "段末缺少句号，建议补全。";
            issue.Priority = 55;
            output.Add(issue);
        }

        private static bool IsHan(char c)
        {
            return c >= '\u4e00' && c <= '\u9fa5';
        }

        private static string DefaultNote(string type, string suggestion)
        {
            if (type == Typo) return "错别字，建议改为“" + suggestion + "”。";
            if (type == Lex) return "用词有误，建议改为“" + suggestion + "”。";
            if (type == Style) return "可考虑改为“" + suggestion + "”。";
            if (type == Redundant) return "语义重复，建议改为“" + suggestion + "”。";
            if (type == Punctuation) return "标点建议：改为“" + suggestion + "”。";
            return "建议改为“" + suggestion + "”。";
        }

        private static List<Issue> ScanNameConsistency(List<string> paragraphs)
        {
            List<Issue> output = new List<Issue>();
            Dictionary<string, string> candidates = new Dictionary<string, string>(StringComparer.Ordinal);

            // Build possible dotted-name forms around each middle dot. This deliberately
            // uses short windows instead of one greedy regex, because Chinese names have
            // no spaces and a greedy match swallows the surrounding sentence.
            for (int i = 0; i < paragraphs.Count; i++)
            {
                string text = paragraphs[i] ?? string.Empty;
                for (int d = 0; d < text.Length; d++)
                {
                    if (text[d] != '·') continue;
                    for (int leftLen = 2; leftLen <= 4; leftLen++)
                    {
                        int leftStart = d - leftLen;
                        if (leftStart < 0) break;
                        bool ok = true;
                        for (int k = leftStart; k < d; k++) if (!IsNameChar(text[k])) { ok = false; break; }
                        if (!ok) break;
                        string left = text.Substring(leftStart, leftLen);
                        for (int rightLen = 2; rightLen <= 6; rightLen++)
                        {
                            int rightEnd = d + 1 + rightLen;
                            if (rightEnd > text.Length) break;
                            bool rightOk = true;
                            for (int k = d + 1; k < rightEnd; k++) if (!IsNameChar(text[k])) { rightOk = false; break; }
                            if (!rightOk) break;
                            string right = text.Substring(d + 1, rightLen);
                            string compact = left + right;
                            if (compact.Length < 4) continue;
                            string dotted = left + "·" + right;
                            string old;
                            if (!candidates.TryGetValue(compact, out old) || dotted.Length > old.Length)
                                candidates[compact] = dotted;
                        }
                    }
                }
            }

            foreach (KeyValuePair<string, string> pair in candidates)
            {
                string compact = pair.Key;
                string full = pair.Value;
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    string text = paragraphs[i] ?? string.Empty;
                    int index = 0;
                    while (index < text.Length)
                    {
                        int hit = text.IndexOf(compact, index, StringComparison.Ordinal);
                        if (hit < 0) break;
                        Issue issue = new Issue();
                        issue.Paragraph = i;
                        issue.Start = hit;
                        issue.Length = compact.Length;
                        issue.Type = Name;
                        issue.Original = compact;
                        issue.Suggestion = full;
                        issue.Note = "文中“" + full + "”带有间隔号，此处写法不一致，建议统一。";
                        issue.Priority = 95;
                        output.Add(issue);
                        index = hit + compact.Length;
                    }
                }
            }
            return output;
        }

        private static bool IsNameChar(char c)
        {
            return (c >= '\u4e00' && c <= '\u9fa5') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }

        private static List<Issue> Resolve(List<Issue> candidates)
        {
            if (candidates.Count == 0) return new List<Issue>();

            // Highest-confidence rules win; then earlier and longer spans.
            candidates.Sort(delegate(Issue a, Issue b)
            {
                int c = b.Priority.CompareTo(a.Priority);
                if (c != 0) return c;
                c = a.Paragraph.CompareTo(b.Paragraph);
                if (c != 0) return c;
                c = a.Start.CompareTo(b.Start);
                if (c != 0) return c;
                return b.Length.CompareTo(a.Length);
            });

            List<Issue> chosen = new List<Issue>();
            foreach (Issue candidate in candidates)
            {
                bool overlap = false;
                for (int i = 0; i < chosen.Count; i++)
                {
                    Issue other = chosen[i];
                    if (other.Paragraph != candidate.Paragraph) continue;
                    if (candidate.Start < other.Start + other.Length && other.Start < candidate.Start + candidate.Length)
                    {
                        overlap = true;
                        break;
                    }
                }
                if (!overlap) chosen.Add(candidate);
                if (chosen.Count >= 5000) break;
            }

            chosen.Sort(delegate(Issue a, Issue b)
            {
                int c = a.Paragraph.CompareTo(b.Paragraph);
                if (c != 0) return c;
                c = a.Start.CompareTo(b.Start);
                if (c != 0) return c;
                return a.Length.CompareTo(b.Length);
            });
            return chosen;
        }

        private static readonly Regex DuplicateCharRegex = new Regex(
            @"(的|是|我|你|他|她|它|就|都|不|很|也|又|还|才|只)\1",
            RegexOptions.None);

        private static List<PhraseRule> BuildPhraseRules()
        {
            List<PhraseRule> list = new List<PhraseRule>();
            string[] lines = PhraseData.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string raw in lines)
            {
                string line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#")) continue;
                string[] parts = line.Split('|');
                if (parts.Length < 3) continue;
                PhraseRule rule = new PhraseRule();
                rule.Wrong = parts[0].Trim();
                rule.Right = parts[1].Trim();
                rule.Type = parts[2].Trim();
                rule.Note = parts.Length >= 4 ? parts[3].Trim() : string.Empty;
                if (rule.Wrong.Length > 0 && rule.Right.Length > 0 && rule.Wrong != rule.Right) list.Add(rule);
            }
            return list;
        }

        private static RegexRule Rule(string pattern, string replacement, string type, string note, int priority)
        {
            RegexRule rule = new RegexRule();
            rule.Pattern = new Regex(pattern, RegexOptions.None);
            rule.Replacement = replacement;
            rule.Type = type;
            rule.Note = note;
            rule.Priority = priority;
            return rule;
        }

        private static RegexRule Rule(string pattern, string replacement, string type, string note, int priority, bool punctuationOnly, bool advancedOnly, bool styleOnly)
        {
            RegexRule rule = Rule(pattern, replacement, type, note, priority);
            rule.PunctuationOnly = punctuationOnly;
            rule.AdvancedOnly = advancedOnly;
            rule.StyleOnly = styleOnly;
            return rule;
        }

        private static List<RegexRule> BuildContextRules()
        {
            List<RegexRule> list = new List<RegexRule>();

            // High-confidence contextual typos.
            list.Add(Rule(@"拷到(?=(?:车|门|窗|栏杆|柱|床|椅|架|管|墙|桌|扶手|暖气|铁门|栅栏|把手))", "铐到", Typo,
                "用手铐固定人时应用“铐”。", 100));
            list.Add(Rule(@"则装在", "则撞在", Typo, "此处应为“撞在”。", 100));
            list.Add(Rule(@"呲牙咧嘴", "龇牙咧嘴", Typo, "规范写法为“龇牙咧嘴”。", 100));
            list.Add(Rule(@"((?:坐|站|躺|趴|醒|爬|跳|冲|走|跑|转|抬|低|直|弯|蹲|跪|靠|躲|藏|探|伸|缩|退|停|愣|怔|笑|哭|叹|松|缓|睁|闭|张|合)(?:起来|过来|下去|出来|进来|回来))“", "$1，“", Punctuation,
                "引语前缺少逗号，建议补上。", 78, true, false, false));

            // 在 / 再
            list.Add(Rule(@"在见(?=[。！？，、…”’]|$)", "再见", Typo, "告别时用“再见”。", 100));
            list.Add(Rule(@"在接再厉", "再接再厉", Typo, "成语为“再接再厉”。", 100));
            list.Add(Rule(@"在次(?=(?:说|问|看|试|想|做|来|等|给|找|检查|确认|提醒|警告|强调|解释|重复|尝试|发生|出现|进行))", "再次", Typo, "表示又一次用“再次”。", 95));

            // 以 / 已 / 己
            list.Add(Rule(@"以经", "已经", Typo, "“已经”表示时间过去。", 100));
            list.Add(Rule(@"已后", "以后", Typo, "“以后”表示时间之后。", 100));
            list.Add(Rule(@"已前", "以前", Typo, "“以前”表示时间之前。", 100));
            list.Add(Rule(@"已便", "以便", Typo, "“以便”表示目的。", 100));
            list.Add(Rule(@"已及", "以及", Typo, "“以及”表示并列。", 100));
            list.Add(Rule(@"已免", "以免", Typo, "“以免”表示避免。", 95));
            list.Add(Rule(@"已防", "以防", Typo, "“以防”表示预防。", 95));
            list.Add(Rule(@"已外", "以外", Typo, "“以外”表示范围之外。", 95));
            list.Add(Rule(@"已内", "以内", Typo, "“以内”表示范围之内。", 95));
            list.Add(Rule(@"已此", "以此", Typo, "“以此”表示凭借这个。", 95));
            list.Add(Rule(@"自已", "自己", Typo, "“自己”指自身。", 100));
            list.Add(Rule(@"己经", "已经", Typo, "“已经”表示时间过去。", 100));
            list.Add(Rule(@"已方", "己方", Typo, "“己方”指自己这一方。", 95));
            list.Add(Rule(@"已任", "己任", Typo, "“己任”指自己的责任。", 95));

            // 因 / 应
            list.Add(Rule(@"因该", "应该", Typo, "“应该”表示理应如此。", 100));
            list.Add(Rule(@"因当", "应当", Typo, "“应当”表示应该。", 100));

            // 即 / 既
            list.Add(Rule(@"即然", "既然", Typo, "“既然”表示前提。", 100));
            list.Add(Rule(@"既使", "即使", Typo, "“即使”表示假设让步。", 100));
            list.Add(Rule(@"既便", "即便", Typo, "“即便”表示即使。", 100));

            // 必需 / 必须
            list.Add(Rule(@"必需(?=(?:去|做|说|得|有|要|坚持|努力|注意|小心|遵守|完成|参加|及时|认真|亲自|先|再|多|少|好|快|慢))", "必须", Lex,
                "表示“一定要”时用“必须”。", 95));
            list.Add(Rule(@"必须品", "必需品", Typo, "“必需品”指不可缺少的物品。", 100));

            // 反应 / 反映
            list.Add(Rule(@"反应(?=(?:情况|问题|意见|诉求|民意|现实|本质|客观|真实|上来|给|到))", "反映", Lex,
                "“反映”指把情况告诉上级或客观呈现。", 95));
            list.Add(Rule(@"反映(?=(?:速度|时间|能力|堆|釜|器|物|式|迟钝|敏捷|迅速|过来))", "反应", Lex,
                "“反应”指机体或事物受刺激后的活动。", 90));

            // 截止 / 截至
            list.Add(Rule(@"截止(?=(?:目前|现在|今日|今天|当前|此刻|到目前|到如今))", "截至", Lex,
                "“截至”表示到某时间为止。", 95));
            list.Add(Rule(@"截止到", "截至", Lex, "“截至”后直接跟时间，不必加“到”。", 90));
            list.Add(Rule(@"截至到", "截至", Lex, "“截至”后不必加“到”。", 100));
            list.Add(Rule(@"截至(?=(?:日期|时间|日|期限))", "截止", Lex, "表示最后期限用“截止”。", 95));

            // 其它 / 交待 等用词建议
            list.Add(Rule(@"其它", "其他", Style, "现代汉语通常写作“其他”。", 60, false, false, true));
            list.Add(Rule(@"交待", "交代", Style, "规范写法多用“交代”。", 55, false, false, true));
            list.Add(Rule(@"身分", "身份", Style, "规范写法多用“身份”。", 55, false, false, true));
            list.Add(Rule(@"想像", "想象", Style, "规范写法多用“想象”。", 55, false, false, true));

            // 过渡 / 过度
            list.Add(Rule(@"过度(?=(?:时期|阶段|政府|地带|季节|形式|办法|措施|方案|政策|社会))", "过渡", Lex,
                "“过渡”指从一个阶段转到另一个阶段；“过度”指超过限度。", 95));
            list.Add(Rule(@"过渡(?=(?:紧张|兴奋|疲劳|劳累|消费|开发|依赖|自信|悲观|乐观|用眼|用药|饮酒|饮食|肥胖|消瘦|焦虑|恐慌|包装))", "过度", Lex,
                "“过度”指超过适当限度。", 95));

            // 侦查 / 侦察
            list.Add(Rule(@"侦察(?=(?:案件|线索|破案|取证))", "侦查", Lex, "公安机关办理案件用“侦查”。", 95));
            list.Add(Rule(@"侦查(?=(?:兵|连|排|机|卫星|设备))", "侦察", Lex, "军事侦察用“侦察”。", 95));

            // 勾通 / 汇合
            list.Add(Rule(@"勾通(?=(?:感情|思想|意见|交流|渠道|能力|双方|一下|起来|彼此))", "沟通", Lex,
                "“沟通”指交流；“勾通”含勾结义。", 95));
            list.Add(Rule(@"汇合(?=(?:队伍|人员|部队|朋友|同事|主力|大军))", "会合", Lex, "人会合，水汇合。", 90));

            // 的 / 地 / 得（疑似规则）
            string adv = @"慢慢|飞快|迅速|悄悄|静静|轻轻|缓缓|渐渐|深深|狠狠|默默|偷偷|匆匆|急忙|连忙|大声|小声|轻声|用力|拼命|仔细|认真|小心|随意|胡乱|随便|特意|故意|不断|不停|逐步|逐渐|纷纷|一一|互相|亲自|独自|一起|高声|低声|沉声|冷声|柔声|厉声|猛然|骤然|霍然|悄然|蓦然|陡然|猛地|忽地|倏地|努力|刻苦|勤奋|用功|专心|耐心|细心|热心|积极|主动|自觉|严格|热烈|深刻|充分|高度|快速|缓慢|轻松|愉快|开心|难过|生气|愤怒|惊讶|疑惑|无奈|坚决|毅然|果断|从容|镇定|慌张|匆忙|悠闲|懒散";
            string verb = @"走|跑|说|问|答|看|望|瞧|瞅|盯|听|想|念|读|写|画|唱|跳|笑|哭|喊|叫|吼|嚷|站|坐|躺|趴|飞|落|流|转|拿|放|推|拉|抓|握|挥|摆|抖|颤|冲|扑|追|逃|躲|藏|停|等|跟|扶|抱|搂|牵|拽|拖|抬|举|扔|丢|摔|砸|踢|打|拍|摸|抚|擦|洗|刷|扫|捡|挑|扛|背|提|捧|端|递|接|关|开|锁|插|拔|按|敲|弹|吹|吃|喝|睡|醒|起|进|出|回|来|去|点|摇|翻|钻|穿|脱|戴|挂|贴|撕|折|剪|切|削|砍|劈|刺|扎|射|骂|夸|劝|催|逼|求|告|诉|讲|谈|聊|商|议|吸|吐|呼|叹|咳嗽|喘气|呼吸|颤抖|点头|摇头|挥手|转身|低头|抬头|皱眉|咬牙|跺脚|迈步|后退|前进|靠近|离开|出现|消失|回答|解释|介绍|询问|叮嘱|吩咐|提醒|警告|责备|抱怨|叹息";
            list.Add(Rule(@"(" + adv + @")的(?=(" + verb + @"))", "$1地", Lex,
                "状语修饰动词，应用“地”。（疑似）", 66, false, true, false));

            string degree = @"很|非常|太|真|极|十分|特别|更|最|不|挺|够";
            string complementVerb = @"跑|走|说|做|写|画|唱|跳|笑|哭|打|干|学|吃|睡|长|过|活|想|看|听|讲|谈|飞|游|表演|完成|处理|解释|回答|准备|收拾|打扫|打扮|生活|工作|学习|比赛|发挥|表现|过";
            list.Add(Rule(@"(" + complementVerb + @")的(?=(" + degree + @"))", "$1得", Lex,
                "程度补语前应用“得”。（疑似）", 66, false, true, false));
            list.Add(Rule(@"(" + complementVerb + @")地(?=(" + degree + @"))", "$1得", Lex,
                "程度补语前应用“得”。（疑似）", 66, false, true, false));

            string complementAdj = @"快|慢|好|坏|对|错|早|晚|远|近|高|低|大|小|多|少|漂亮|难看|干净|整齐|清楚|明白|开心|难过|累|饿|困|疼|冷|热|香|甜|苦|辣|响|亮|暗|拥挤|热闹|安静|去|丑";
            string simpleVerb = @"跑|走|说|做|写|唱|跳|笑|哭|打|干|学|吃|睡|长|过|活|想|看|听|讲|谈|飞|游|穿|洗|扫|收拾|准备|完成|解释|回答";
            list.Add(Rule(@"(" + simpleVerb + @")的(?=(" + complementAdj + @"))", "$1得$2", Lex,
                "程度补语前应用“得”。（疑似）", 62, false, true, false));
            list.Add(Rule(@"(" + simpleVerb + @")地(?=(" + complementAdj + @"))", "$1得$2", Lex,
                "程度补语前应用“得”。（疑似）", 62, false, true, false));

            string adj = @"美丽|漂亮|高大|矮小|宽阔|狭窄|明亮|黑暗|干净|肮脏|整齐|凌乱|温暖|寒冷|炎热|凉爽|安静|喧闹|热闹|冷清|繁华|荒凉|古老|年轻|崭新|破旧|精致|粗糙|柔软|坚硬|光滑|湿润|干燥|香甜|苦涩|清脆|低沉|洪亮|微弱|浓烈|清淡|鲜艳|暗淡|雪白|漆黑|通红|碧绿|金黄|蔚蓝";
            string noun = @"天空|阳光|月光|灯光|房间|屋子|街道|城市|村庄|山|河|海|湖|树|花|草|脸|眼睛|头发|衣服|裙子|声音|笑容|表情|目光|背影|风景|景色|建筑|房子|桌子|椅子|道路|大地";
            list.Add(Rule(@"(" + adj + @")地(?=(" + noun + @"))", "$1的$2", Lex,
                "定语修饰名词，应用“的”。（疑似）", 60, false, true, false));

            // Punctuation only rules.
            list.Add(Rule(@"([，。！？；：、])\1+", "$1", Punctuation, "重复标点，保留一个即可。", 60, true, false, false));
            list.Add(Rule(@"([，。！？；：、])( +)", "$1", Punctuation, "中文标点后不宜加空格。", 55, true, false, false));
            list.Add(Rule(@"(?<=[\u4e00-\u9fa5]),(?=[\u4e00-\u9fa5])", "，", Punctuation, "中文语境应使用全角逗号。", 58, true, false, false));
            list.Add(Rule(@"(?<=[\u4e00-\u9fa5])\.(?=[\u4e00-\u9fa5])", "。", Punctuation, "中文语境应使用全角句号。", 58, true, false, false));
            list.Add(Rule(@"(?<=[\u4e00-\u9fa5])!(?=[\u4e00-\u9fa5])", "！", Punctuation, "中文语境应使用全角叹号。", 58, true, false, false));
            list.Add(Rule(@"(?<=[\u4e00-\u9fa5])\?(?=[\u4e00-\u9fa5])", "？", Punctuation, "中文语境应使用全角问号。", 58, true, false, false));
            list.Add(Rule(@"(?<=[\u4e00-\u9fa5]);(?=[\u4e00-\u9fa5])", "；", Punctuation, "中文语境应使用全角分号。", 58, true, false, false));
            list.Add(Rule(@"(?<=[\u4e00-\u9fa5]):(?=[\u4e00-\u9fa5])", "：", Punctuation, "中文语境应使用全角冒号。", 58, true, false, false));
            list.Add(Rule(@"(?<=[\u4e00-\u9fa5])\.{2,}(?=[\u4e00-\u9fa5]|$)", "……", Punctuation, "中文省略号应写作“……”（六点）。", 62, true, false, false));
            list.Add(Rule(@"(?<=[\u4e00-\u9fa5])[!?;,](?=$)", "。", Punctuation, "段末应使用中文标点。", 55, true, false, false));

            // A few redundancy rules.
            list.Add(Rule(@"凯旋而归", "凯旋", Redundant, "“凯旋”已含“归来”之意。", 100));
            list.Add(Rule(@"忍俊不禁地笑", "忍俊不禁", Redundant, "“忍俊不禁”已含“笑”的意思。", 100));
            list.Add(Rule(@"(?<p>(?:大约|约|将近|接近)[^。！？，]{1,10})左右", "${p}", Redundant, "“约”与“左右”语义重复，保留一个。", 90));
            list.Add(Rule(@"(?<p>(?:至少|最少))([^。！？，]{1,10})以上", "${p}${2}", Redundant, "“至少”与“以上”语义重复。", 90));
            list.Add(Rule(@"(?<p>(?:最多|至多))([^。！？，]{1,10})以下", "${p}${2}", Redundant, "“最多”与“以下”语义重复。", 90));
            list.Add(Rule(@"(?<p>超过)([^。！？，]{1,10})以上", "${p}${2}", Redundant, "“超过”与“以上”语义重复。", 90));
            list.Add(Rule(@"首先第一", "首先", Redundant, "“首先”与“第一”语义重复。", 90));
            list.Add(Rule(@"目前的当务之急", "当务之急", Redundant, "“当务之急”已含“目前”之意。", 90));
            list.Add(Rule(@"开始启动", "启动", Redundant, "“启动”已含“开始”之意。", 90));
            list.Add(Rule(@"彻底根治", "根治", Redundant, "“根治”已含“彻底”之意。", 90));
            list.Add(Rule(@"一致公认", "公认", Redundant, "“公认”已含“一致”之意。", 90));
            list.Add(Rule(@"突然恍然大悟", "恍然大悟", Redundant, "“恍然大悟”已含“突然”之意。", 90));
            list.Add(Rule(@"再次重申", "重申", Redundant, "“重申”已含“再次”之意。", 90));
            list.Add(Rule(@"提前预知", "预知", Redundant, "“预知”已含“提前”之意。", 90));
            list.Add(Rule(@"提前预警", "预警", Redundant, "“预警”已含“提前”之意。", 90));
            list.Add(Rule(@"非法走私", "走私", Redundant, "“走私”本身即含非法义。", 90));
            list.Add(Rule(@"众多莘莘学子", "莘莘学子", Redundant, "“莘莘”已表示众多。", 90));
            list.Add(Rule(@"涉及到", "涉及", Redundant, "“涉及”后不必加“到”。", 85));
            list.Add(Rule(@"波及到", "波及", Redundant, "“波及”后不必加“到”。", 85));
            list.Add(Rule(@"付诸于", "付诸", Redundant, "“付诸”后不必加“于”。", 90));
            list.Add(Rule(@"诉诸于", "诉诸", Redundant, "“诉诸”后不必加“于”。", 90));
            list.Add(Rule(@"来自于", "来自", Redundant, "“来自”后不必加“于”。", 80));
            list.Add(Rule(@"原因是由于", "原因是", Redundant, "“原因是”与“由于”语义重复。", 80));
            list.Add(Rule(@"原因是因为", "原因是", Redundant, "“原因是”与“因为”语义重复。", 80));
            list.Add(Rule(@"切忌不要", "切忌", Redundant, "“切忌”已含“不要”之意。", 90));
            list.Add(Rule(@"避免不再(?=(?:发生|出现))", "避免再次", Redundant, "“避免不再发生”意思相反，应改为“避免再次发生”。", 95));
            list.Add(Rule(@"防止不再(?=(?:发生|出现))", "防止再次", Redundant, "“防止不再发生”意思相反，应改为“防止再次发生”。", 95));

            return list;
        }

        private const string PhraseData = @"
# 常见成语/词语误写
按耐不住|按捺不住|错别字
迫不急待|迫不及待|错别字
一如继往|一如既往|错别字
一如即往|一如既往|错别字
谈笑风声|谈笑风生|错别字
再所难免|在所难免|错别字
走头无路|走投无路|错别字
变本加利|变本加厉|错别字
伶牙利齿|伶牙俐齿|错别字
悬梁刺骨|悬梁刺股|错别字
不径而走|不胫而走|错别字
默守成规|墨守成规|错别字
一股作气|一鼓作气|错别字
萎糜不振|萎靡不振|错别字
穿流不息|川流不息|错别字
委屈求全|委曲求全|错别字
直接了当|直截了当|错别字
别出新裁|别出心裁|错别字
陈词烂调|陈词滥调|错别字
鬼鬼崇崇|鬼鬼祟祟|错别字
汗流夹背|汗流浃背|错别字
轰堂大笑|哄堂大笑|错别字
既往不究|既往不咎|错别字
娇揉造作|矫揉造作|错别字
竭泽而鱼|竭泽而渔|错别字
金榜提名|金榜题名|错别字
精兵减政|精兵简政|错别字
精神焕散|精神涣散|错别字
苦心孤旨|苦心孤诣|错别字
滥芋充数|滥竽充数|错别字
鳞次节比|鳞次栉比|错别字
满腹经论|满腹经纶|错别字
迷天大谎|弥天大谎|错别字
名列前矛|名列前茅|错别字
明火执杖|明火执仗|错别字
磨拳擦掌|摩拳擦掌|错别字
莫明其妙|莫名其妙|错别字
弄巧成绌|弄巧成拙|错别字
披星带月|披星戴月|错别字
破斧沉舟|破釜沉舟|错别字
千锤百练|千锤百炼|错别字
前扑后继|前仆后继|错别字
轻歌慢舞|轻歌曼舞|错别字
情不自尽|情不自禁|错别字
趋之若骛|趋之若鹜|错别字
人才倍出|人才辈出|错别字
如火如茶|如火如荼|错别字
入不付出|入不敷出|错别字
弱不经风|弱不禁风|错别字
三翻五次|三番五次|错别字
杀一敬百|杀一儆百|错别字
山青水秀|山清水秀|错别字
少安勿躁|少安毋躁|错别字
声名雀起|声名鹊起|错别字
世外桃园|世外桃源|错别字
首曲一指|首屈一指|错别字
水乳交溶|水乳交融|错别字
死皮癞脸|死皮赖脸|错别字
贪脏枉法|贪赃枉法|错别字
天翻地复|天翻地覆|错别字
甜言密语|甜言蜜语|错别字
挑拔离间|挑拨离间|错别字
同仇敌慨|同仇敌忾|错别字
图穷匕现|图穷匕见|错别字
推心至腹|推心置腹|错别字
万古常青|万古长青|错别字
枉废心机|枉费心机|错别字
望而怯步|望而却步|错别字
微言大意|微言大义|错别字
无动于中|无动于衷|错别字
无精打彩|无精打采|错别字
无是生非|无事生非|错别字
五彩斑澜|五彩斑斓|错别字
喜出忘外|喜出望外|错别字
相辅相承|相辅相成|错别字
相形见拙|相形见绌|错别字
消声匿迹|销声匿迹|错别字
心浮气燥|心浮气躁|错别字
心恢意冷|心灰意冷|错别字
兴高彩烈|兴高采烈|错别字
凶神恶刹|凶神恶煞|错别字
雪中送碳|雪中送炭|错别字
寻人启示|寻人启事|错别字
招聘启示|招聘启事|错别字
鸦鹊无声|鸦雀无声|错别字
掩旗息鼓|偃旗息鼓|错别字
一愁莫展|一筹莫展|错别字
一诺千斤|一诺千金|错别字
一泄千里|一泻千里|错别字
义无返顾|义无反顾|错别字
引经据点|引经据典|错别字
应接不瑕|应接不暇|错别字
忧心冲冲|忧心忡忡|错别字
与日具增|与日俱增|错别字
欲盖弥章|欲盖弥彰|错别字
原形必露|原形毕露|错别字
再接再励|再接再厉|错别字
责无旁代|责无旁贷|错别字
震聋发聩|振聋发聩|错别字
支言片语|只言片语|错别字
中流抵柱|中流砥柱|错别字
专心至志|专心致志|错别字
自抱自弃|自暴自弃|错别字
坐想其成|坐享其成|错别字
慢不经心|漫不经心|错别字
美仑美奂|美轮美奂|错别字
梦昧以求|梦寐以求|错别字
磨肩接踵|摩肩接踵|错别字
旁证博引|旁征博引|错别字
迫在眉捷|迫在眉睫|错别字
气势凶凶|气势汹汹|错别字
千均一发|千钧一发|错别字
前踞后恭|前倨后恭|错别字
黔驴计穷|黔驴技穷|错别字
清淅可见|清晰可见|错别字
情有独衷|情有独钟|错别字
穷困缭倒|穷困潦倒|错别字
曲指可数|屈指可数|错别字
全神灌注|全神贯注|错别字
热血沸滕|热血沸腾|错别字
如法泡制|如法炮制|错别字
如雷灌耳|如雷贯耳|错别字
神色张惶|神色张皇|错别字
生死悠关|生死攸关|错别字
声名狼籍|声名狼藉|错别字
拾人牙惠|拾人牙慧|错别字
势不两力|势不两立|错别字
瘦骨鳞峋|瘦骨嶙峋|错别字
首当其充|首当其冲|错别字
书声朗朗|书声琅琅|错别字
束手无册|束手无策|错别字
水泻不通|水泄不通|错别字
死心踏地|死心塌地|错别字
俗不可奈|俗不可耐|错别字
损兵拆将|损兵折将|错别字
天崖海角|天涯海角|错别字
同流合乌|同流合污|错别字
突如奇来|突如其来|错别字
推波助蓝|推波助澜|错别字
脱颍而出|脱颖而出|错别字
歪风斜气|歪风邪气|错别字
完壁归赵|完璧归赵|错别字
万籁具寂|万籁俱寂|错别字
妄费心机|枉费心机|错别字
微不足到|微不足道|错别字
为虎作怅|为虎作伥|错别字
五体头地|五体投地|错别字
物及必反|物极必反|错别字
喜笑眼开|喜笑颜开|错别字
闲情逸志|闲情逸致|错别字
想往已久|向往已久|错别字
笑逐眼开|笑逐颜开|错别字
心狠手棘|心狠手辣|错别字
信口开合|信口开河|错别字
兴师动重|兴师动众|错别字
形消骨立|形销骨立|错别字
兄有成竹|胸有成竹|错别字
虚座以待|虚左以待|错别字
摇摇欲堕|摇摇欲坠|错别字
一脉相成|一脉相承|错别字
怨天由人|怨天尤人|错别字
张灯结采|张灯结彩|错别字
真知卓见|真知灼见|错别字
直言不诲|直言不讳|错别字
指手划脚|指手画脚|错别字
趾高气昂|趾高气扬|错别字
众口烁金|众口铄金|错别字
自命不番|自命不凡|错别字
自食其利|自食其力|错别字
慢条斯礼|慢条斯理|错别字
毛骨耸然|毛骨悚然|错别字
迷或不解|迷惑不解|错别字
面面具到|面面俱到|错别字
明辩是非|明辨是非|错别字
目不瑕接|目不暇接|错别字
脑羞成怒|恼羞成怒|错别字
披肝厉胆|披肝沥胆|错别字
潜移墨化|潜移默化|错别字
全神惯注|全神贯注|错别字
三令五伸|三令五申|错别字
肆无忌弹|肆无忌惮|错别字
惮精竭虑|殚精竭虑|错别字
心弛神往|心驰神往|错别字
张驰有度|张弛有度|错别字
弛名中外|驰名中外|错别字
病入膏盲|病入膏肓|错别字
脍灸人口|脍炙人口|错别字
好高鹜远|好高骛远|错别字
心无旁鹜|心无旁骛|错别字
绝然不同|截然不同|错别字
大相经庭|大相径庭|错别字
竟直|径直|错别字
峻工|竣工|错别字
严竣|严峻|错别字
险竣|险峻|错别字
疏竣|疏浚|错别字
崇山竣岭|崇山峻岭|错别字
暇疵|瑕疵|错别字
洁白无暇|洁白无瑕|错别字
完美无暇|完美无瑕|错别字
暇不掩瑜|瑕不掩瑜|错别字
苍海桑田|沧海桑田|错别字
苍海一粟|沧海一粟|错别字
磋跎|蹉跎|错别字
旁徨|彷徨|错别字
俳徊|徘徊|错别字
倘佯|徜徉|错别字
踉呛|踉跄|错别字
咧趄|趔趄|错别字
盘跚|蹒跚|错别字
蜿延|蜿蜒|错别字
委迤|逶迤|错别字
氲氤|氤氲|错别字
缱卷|缱绻|错别字
婆挲|婆娑|错别字
斑剥|斑驳|错别字
斑烂|斑斓|错别字
干旋|斡旋|错别字
懈逅|邂逅|错别字
技俩|伎俩|错别字
老俩口|老两口|错别字
小俩口|小两口|错别字
装贞|装帧|错别字
装璜|装潢|错别字
屏幛|屏障|错别字
层峦叠障|层峦叠嶂|错别字
书藉|书籍|错别字
蕴籍|蕴藉|错别字
藉贯|籍贯|错别字
国藉|国籍|错别字
户藉|户籍|错别字
古藉|古籍|错别字
凭藉|凭借|用词建议
藉口|借口|用词建议
藉故|借故|用词建议
藉此|借此|用词建议
包函|包涵|错别字
涵概|涵盖|错别字
函盖|涵盖|错别字
函养|涵养|错别字
涵授|函授|错别字
慨括|概括|错别字
感概|感慨|错别字
慷概|慷慨|错别字
大慨|大概|错别字
气慨|气概|错别字
良霄|良宵|错别字
春霄|春宵|错别字
夜霄|夜宵|错别字
通霄达旦|通宵达旦|错别字
直冲云宵|直冲云霄|错别字
宵壤之别|霄壤之别|错别字
九宵云外|九霄云外|错别字
气冲宵汉|气冲霄汉|错别字
毗临|毗邻|错别字
临居|邻居|错别字
临座|邻座|错别字
临国|邻国|错别字
临邦|邻邦|错别字
挤身|跻身|错别字
闲熟|娴熟|错别字
造旨|造诣|错别字
瞻养|赡养|错别字
赡仰|瞻仰|错别字
赡前顾后|瞻前顾后|错别字
赊帐|赊账|错别字
十恶不郝|十恶不赦|错别字
郝郝有名|赫赫有名|错别字
道欠|道歉|错别字
欠意|歉意|错别字
抱欠|抱歉|错别字
报歉|抱歉|错别字
报怨|抱怨|错别字
抱酬|报酬|错别字
抱答|报答|错别字
抱效|报效|错别字
抱国|报国|错别字
抱仇|报仇|错别字
报愧|抱愧|错别字
暴炸|爆炸|错别字
暴发力|爆发力|错别字
火山暴发|火山爆发|错别字
爆怒|暴怒|错别字
爆躁|暴躁|错别字
爆力|暴力|错别字
爆利|暴利|错别字
爆涨|暴涨|错别字
爆跌|暴跌|错别字
暴满|爆满|错别字
暴笑|爆笑|错别字
暴竹|爆竹|错别字
暴胎|爆胎|错别字
暴光|曝光|错别字
爆光|曝光|错别字
暴棚|爆棚|错别字
振惊|震惊|错别字
振撼|震撼|错别字
振耳欲聋|震耳欲聋|错别字
振怒|震怒|错别字
地振|地震|错别字
震奋|振奋|错别字
震兴|振兴|错别字
震作|振作|错别字
震震有词|振振有词|错别字
震臂高呼|振臂高呼|错别字
震幅|振幅|错别字
蓝球|篮球|错别字
蓝子|篮子|错别字
竹蓝|竹篮|错别字
花蓝|花篮|错别字
摇蓝|摇篮|错别字
篮图|蓝图|错别字
蔚篮|蔚蓝|错别字
篮天|蓝天|错别字
篮色|蓝色|错别字
篮领|蓝领|错别字
篮本|蓝本|错别字
夜暮|夜幕|错别字
幕色|暮色|错别字
幕年|暮年|错别字
幕春|暮春|错别字
朝幕|朝暮|错别字
日幕|日暮|错别字
垂幕|垂暮|错别字
幕霭|暮霭|错别字
幕名|慕名|错别字
羡幕|羡慕|错别字
仰幕|仰慕|错别字
爱幕|爱慕|错别字
敬幕|敬慕|错别字
渴幕|渴慕|错别字
思幕|思慕|错别字
幕穴|墓穴|错别字
幕碑|墓碑|错别字
幕地|墓地|错别字
扫幕|扫墓|错别字
坟幕|坟墓|错别字
公幕|公墓|错别字
墓布|幕布|错别字
墓僚|幕僚|错别字
墓后|幕后|错别字
序暮|序幕|错别字
帷墓|帷幕|错别字
银墓|银幕|错别字
屏墓|屏幕|错别字
字墓|字幕|错别字
荧墓|荧幕|错别字
墙璧|墙壁|错别字
璧纸|壁纸|错别字
璧橱|壁橱|错别字
璧画|壁画|错别字
璧虎|壁虎|错别字
峭璧|峭壁|错别字
隔璧|隔壁|错别字
碰璧|碰壁|错别字
白壁无瑕|白璧无瑕|错别字
珠联壁合|珠联璧合|错别字
中西合壁|中西合璧|错别字
和氏壁|和氏璧|错别字
竟争|竞争|错别字
竟赛|竞赛|错别字
竟选|竞选|错别字
竟技|竞技|错别字
竟标|竞标|错别字
竟价|竞价|错别字
竟相|竞相|错别字
竟逐|竞逐|错别字
必竟|毕竟|错别字
究竞|究竟|错别字
神彩|神采|错别字
光采|光彩|错别字
风彩|风采|错别字
精采|精彩|错别字
喝采|喝彩|错别字
采排|彩排|错别字
剪采|剪彩|错别字
采虹|彩虹|错别字
丰富多采|丰富多彩|错别字
兴高彩烈|兴高采烈|错别字
无精打彩|无精打采|错别字
理采|理睬|错别字
不理不采|不理不睬|错别字
带眼镜|戴眼镜|错别字
带着帽子|戴着帽子|错别字
带手套|戴手套|错别字
带围巾|戴围巾|错别字
带口罩|戴口罩|错别字
穿带|穿戴|错别字
爱带|爱戴|错别字
带罪立功|戴罪立功|错别字
张冠李带|张冠李戴|错别字
披星带月|披星戴月|错别字
佩带首饰|佩戴首饰|用词建议
佩带项链|佩戴项链|用词建议
佩带戒指|佩戴戒指|用词建议
佩带眼镜|佩戴眼镜|用词建议
坐位|座位|错别字
座落|坐落|错别字
坐右铭|座右铭|错别字
坐次|座次|错别字
坐椅|座椅|错别字
坐钟|座钟|错别字
坐机|座机|错别字
坐驾|座驾|错别字
一坐山|一座山|错别字
一坐城|一座城|错别字
一坐桥|一座桥|错别字
一坐楼|一座楼|错别字
一坐庙|一座庙|错别字
一坐塔|一座塔|错别字
一坐雕像|一座雕像|错别字
一坐钟|一座钟|错别字
一坐工厂|一座工厂|错别字
一坐学校|一座学校|错别字
一坐医院|一座医院|错别字
一坐车站|一座车站|错别字
一坐宫殿|一座宫殿|错别字
一坐房子|一座房子|错别字
一坐建筑|一座建筑|错别字
做在椅子|坐在椅子|错别字
做在沙发|坐在沙发|错别字
做在床上|坐在床上|错别字
做在地上|坐在地上|错别字
做在位置|坐在位置|错别字
做在座位|坐在座位|错别字
做在门口|坐在门口|错别字
做在窗边|坐在窗边|错别字
做在桌前|坐在桌前|错别字
做位|座位|错别字
做落|坐落|错别字
做右铭|座右铭|错别字
作饭|做饭|错别字
作梦|做梦|错别字
作功课|做功课|错别字
作作业|做作业|错别字
作生意|做生意|错别字
作事情|做事情|错别字
作东|做东|错别字
作做|做作|错别字
那怕|哪怕|错别字
那能|哪能|错别字
那敢|哪敢|错别字
那知|哪知|错别字
那料|哪料|错别字
那管|哪管|错别字
那曾想|哪曾想|错别字
在接再厉|再接再厉|错别字
好象|好像|错别字
就象|就像|错别字
象是|像是|错别字
不象|不像|错别字
更象|更像|错别字
最象|最像|错别字
图象|图像|错别字
录象|录像|错别字
影象|影像|错别字
肖象|肖像|错别字
画象|画像|错别字
偶象|偶像|错别字
录相|录像|错别字
摄相|摄像|错别字
照像|照相|错别字
照像机|照相机|错别字
录相机|录像机|错别字
摄相机|摄像机|错别字
相象|相像|错别字
真象|真相|用词建议
假相|假象|用词建议
一昧|一味|错别字
暖昧|暧昧|错别字
暗然|黯然|错别字
暗然失色|黯然失色|错别字
和霭|和蔼|错别字
暮蔼|暮霭|错别字
雾蔼|雾霭|错别字
安祥|安详|错别字
祥细|详细|错别字
端祥|端详|错别字
吉详|吉祥|错别字
慈详|慈祥|错别字
寒喧|寒暄|错别字
暄哗|喧哗|错别字
暄嚣|喧嚣|错别字
渲泄|宣泄|错别字
喧染|渲染|错别字
密诀|秘诀|错别字
秘决|秘诀|错别字
决窍|诀窍|错别字
决别|诀别|错别字
永决|永诀|错别字
要决|要诀|错别字
决择|抉择|错别字
抉别|诀别|错别字
决对|绝对|错别字
绝定|决定|错别字
绝裂|决裂|错别字
秘蜜|秘密|错别字
甜密|甜蜜|错别字
亲蜜|亲密|错别字
密月|蜜月|错别字
甜言密语|甜言蜜语|错别字
口密腹剑|口蜜腹剑|错别字
密方|秘方|错别字
密而不宣|秘而不宣|用词建议
蜇伏|蛰伏|错别字
惊蜇|惊蛰|错别字
床第|床笫|错别字
汨汨|汩汩|错别字
针贬|针砭|错别字
针贬时弊|针砭时弊|错别字
编篡|编纂|错别字
修茸|修葺|错别字
肆业|肄业|错别字
膺品|赝品|错别字
痉孪|痉挛|错别字
分岐|分歧|错别字
岐路|歧路|错别字
岐视|歧视|错别字
岐途|歧途|错别字
撑杆跳|撑竿跳|错别字
";
    }
}
