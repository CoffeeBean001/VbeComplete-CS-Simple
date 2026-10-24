// 纯判据：不碰 COM、不碰 UI、不碰全局状态 —— 全部可以单独测。
//
// 为什么这些要独立成一个文件：它们是"补全准不准"的全部决定者。
//   混进 COM 或窗口代码里的话，就只能靠真机试，一条 bug 要试十次。

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal static class Rules
    {
        public static bool IsIdentChar(char c) => char.IsLetterOrDigit(c) || c == '_';

        /// <summary>
        /// 取光标左边那个词。caretCol 是 1-based 字符列（VBE 的 GetSelection 给的）。
        /// 返回 (词, 词的起始列 1-based)；取不到返回 ("", caretCol)。
        /// </summary>
        public static string WordBefore(string line, int caretCol, out int wordStartCol)
        {
            wordStartCol = caretCol;
            if (string.IsNullOrEmpty(line)) return "";
            int end = Math.Min(caretCol - 1, line.Length);   // 0-based 排他
            if (end <= 0) return "";
            int i = end - 1;
            if (!IsIdentChar(line[i])) return "";
            int j = i;
            while (j >= 0 && IsIdentChar(line[j])) j--;
            wordStartCol = j + 2;                            // 1-based
            return line.Substring(j + 1, i - j);
        }

        /// <summary>把 [wordStartCol, endCol) 换成 completion（列都是 1-based）。</summary>
        public static string ReplaceWord(string line, int wordStartCol, int endCol, string completion)
        {
            line = line ?? "";
            // 先把 start 夹进行宽，再算 end —— 反过来的话 start 越界时 end 会停在
            //   start 上（也越界），下面 Substring(end) 直接抛 ArgumentOutOfRange。
            //   "实际行比快照行短"真的会发生：撤销 / 别处并发改动 / VBE 自己重排。
            int start = Math.Max(0, Math.Min(wordStartCol - 1, line.Length));
            int end = Math.Max(start, Math.Min(line.Length, endCol - 1));
            return line.Substring(0, start) + completion + line.Substring(end);
        }

        /// <summary>
        /// VBE 重排过整行之后，找回补全词真正的结束下标（0-based 排他）。
        /// 只认【整词】出现（前后都不是标识符字符），多处出现取起点最接近 near0 的；
        /// 找不到就用 fallback0。
        /// 原来用 LastIndexOf 且不校验词边界 —— 行里出现两次同名
        ///   （`Split = Split(x)`）时会取到最后一个，光标就跑到别的词后面去。
        ///   判据与取词结束位置的判定共用同一份语义。
        /// </summary>
        public static int FindCompletionEnd(string actual, string completion, int near0, int fallback0)
        {
            int clamped = Math.Max(0, Math.Min(actual == null ? 0 : actual.Length, fallback0));
            if (string.IsNullOrEmpty(actual) || string.IsNullOrEmpty(completion)) return clamped;
            int L = completion.Length;
            int best = -1;
            int i = 0;
            while (i <= actual.Length - L)
            {
                int hit = actual.IndexOf(completion, i, StringComparison.OrdinalIgnoreCase);
                if (hit < 0) break;
                bool okL = hit == 0 || !IsIdentChar(actual[hit - 1]);
                bool okR = hit + L >= actual.Length || !IsIdentChar(actual[hit + L]);
                if (okL && okR && (best < 0 || Math.Abs(hit - near0) < Math.Abs(best - near0))) best = hit;
                i = hit + L;
            }
            if (best >= 0) return best + L;
            return clamped;
        }

        /// <summary>
        /// VBE 编辑器里一个 Tab 等于几个显示列。出厂值是 4。
        /// <para>从 settings.ini 的 `TAB_WIDTH` 读（1~16）。三个换算函数传
        /// `tabw &lt;= 0` 就表示"用这个全局值"。VBE 的制表符宽度不是 4 时，
        /// 把这个值调成一样的，候选取词和光标落点才对得上。</para>
        /// </summary>
        public static int TabWidth = Settings.Int("TAB_WIDTH", 4, 1, 16);

        /// <summary>
        /// 字符下标 -> VBE 认的"显示列"（1-based）。
        /// VBE 的 SetSelection 列是按 Tab 展开后的格数算的：不做换算，光标
        /// 在有 Tab 或缺省缩进（VBE 用 Tab）的行上会落进词的中间。
        /// </summary>
        public static int DispCol(string text, int charIndex, int tabw = -1)
        {
            if (tabw <= 0) tabw = TabWidth;
            if (string.IsNullOrEmpty(text)) return 1;
            if (charIndex > text.Length) charIndex = text.Length;
            if (charIndex < 0) charIndex = 0;
            int col = 0;
            for (int i = 0; i < charIndex; i++)
            {
                char c = text[i];
                if (c == '\t') col += tabw - (col % tabw);
                else if (IsWide(c)) col += 2;
                else col += 1;
            }
            return col + 1;
        }

        /// <summary>
        /// "第 charCol 个字符之前" -> VBE 认的显示列（1-based）。
        /// `SetCaret` 的列语义唯一实现。DispCol(i) 给的是"字符 i 之前"，
        ///   所以第 charCol 个字符之前 = DispCol(charCol - 1)。差这一格，真机上
        ///   就是"光标永远比你以为的右一格"。
        /// </summary>
        public static int CaretColForCharCol(string text, int charCol, int tabw = -1)
        {
            if (charCol < 1) charCol = 1;
            return DispCol(text, charCol - 1, tabw);
        }

        public static bool IsWide(char c) =>
            (c >= 0x1100 && c <= 0x115F) || (c >= 0x2E80 && c <= 0xA4CF) ||
            (c >= 0xAC00 && c <= 0xD7A3) || (c >= 0xF900 && c <= 0xFAFF) ||
            (c >= 0xFE30 && c <= 0xFE6F) || (c >= 0xFF00 && c <= 0xFF60) ||
            (c >= 0xFFE0 && c <= 0xFFE6);

        /// <summary>
        /// 反向换算：VBE 给的"显示列"（1-based）-> 1-based 字符列。
        /// `CodePane.GetSelection` 返回的列是**显示列**（真机实测：
        ///   `    学生 = Split()` 里 `(` 前是 17 列，按字符数是 15）。把它直接
        ///   当字符列去取词，含中文/全角的行会取错位置甚至取到空词 ——
        ///   症状是"这一行根本不弹候选"。凡是吃 `snap.CaretCol` 的判据，
        ///   都必须先过这一道。
        /// </summary>
        public static int CharColFromDisp(string text, int dispCol, int tabw = -1)
        {
            if (tabw <= 0) tabw = TabWidth;
            if (string.IsNullOrEmpty(text) || dispCol <= 1) return 1;
            int col = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                int w = c == '\t' ? tabw - (col % tabw) : (IsWide(c) ? 2 : 1);
                if (col + w >= dispCol) return i + 1;
                col += w;
            }
            return text.Length + 1;
        }

        // ------------------------------------------------------------------
        // 位置判据：这个位置该不该弹
        // ------------------------------------------------------------------

        /// <summary>
        /// 声明语句的"头"：带 Dim 的常规声明，以及**不带 Dim 的裸声明**。
        ///
        /// 补上裸声明这一支：`Public gCount` / `Private m_arr` / `Global g`
        ///   是模块级变量的常见写法，旧正则只认 `Public Sub` 这种"修饰符 + 关键字"，
        ///   于是 `Public g|` 被判成普通表达式 —— 弹窗照弹，而候选池里刚好有
        ///   它自己（解析时把 `Public g` 当成了已声明的名字），用户看到的就是
        ///   "定义全局变量时提示它自己"。
        /// </summary>
        private static readonly Regex ReDeclHead = new Regex(
            @"^[ \t]*(?:" +
            @"(?:(?:Public|Private|Friend|Global|Static)[ \t]+)*" +
            @"(?:Dim|Const|ReDim|WithEvents|Sub|Function|Property|Declare|Event|Enum|Type)\b" +
            @"|(?:Public|Private|Friend|Global|Static)[ \t]+" +
            @"(?!(?:Sub|Function|Property|Declare|Event|Enum|Type|Const|Dim|ReDim|WithEvents)\b)" +
            @")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// 光标是不是停在"造一个新名字 / 写一个类型"的位置 —— 这类位置一律不弹候选。
        ///
        /// 起语义变了（口径）：
        ///   * `As` 后面【不弹】。那里只可能是类型名，VBE 自己的处理更准，
        ///     我们弹一层上去只会挡住它、打断打字节奏。
        ///   * 裸声明的名字位置【不弹】（见 ReDeclHead 的说明）。
        ///   例外：`Const PI = 3.14` 的等号右边是**表达式**，那里要照常提示。
        /// </summary>
        public static bool InDeclarationPosition(string line, int caretCol)
        {
            if (string.IsNullOrEmpty(line)) return false;
            int end = Math.Max(0, Math.Min(caretCol - 1, line.Length));

            // 只看**光标所在的语句段**：VBA 允许 `Dim x As Long: x = Ms|`，
            //   整行以 Dim 开头不代表光标这段是声明 —— 用整行判会把冒号后面
            //   那条普通赋值也当成"造新名字"，从此在那里永远不弹候选。
            int segStart = StatementStart(line, end);
            string seg = line.Substring(segStart, end - segStart);
            if (!ReDeclHead.IsMatch(seg)) return false;

            // `Const PI = 3.14` 的值位置不是名字位置，也不是类型位置
            if (Regex.IsMatch(seg, @"\bConst\b.*=", RegexOptions.IgnoreCase)) return false;
            return true;
        }

        /// <summary>
        /// 光标所在**语句段**的起点（0-based）。VBA 用 `:` 分隔同一行里的多条语句，
        /// 只认「引号外、注释外」的冒号 —— `s = "a:b"` 里的那个不算分隔符。
        /// </summary>
        public static int StatementStart(string line, int endCol)
        {
            if (string.IsNullOrEmpty(line)) return 0;
            int end = Math.Max(0, Math.Min(endCol, line.Length));
            int start = 0;
            bool inStr = false;
            for (int i = 0; i < end; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    // VBA 的字符串转义是双写引号（`""`），跳过去别误判成"关闭"
                    if (inStr && i + 1 < end && line[i + 1] == '"') { i++; continue; }
                    inStr = !inStr;
                    continue;
                }
                if (inStr) continue;
                if (c == '\'') break;        // 注释从这里开始，后面没有语句了
                if (c == ':') start = i + 1;
            }
            return start;
        }

        /// <summary>
        /// 光标前那个词是不是紧跟在 `As` / `New` 这类"后面只可能是类型名"的关键字后面。
        ///
        /// 为什么单独一条、不并进 <see cref="InDeclarationPosition"/>：
        ///   那条只认 `Dim` / `Const` / `ReDim` / `WithEvents` **开头的整行**；
        ///   而 `Set c = New |`、`dic.Add 1, New |`、`For i As |`、过程参数
        ///   `Sub Foo(a As |)` 的行首都不是声明关键字 —— 不判就会照弹。
        ///   口径`New` 后面跟 `As` 一样，不弹。
        ///
        ///   wordStartCol 是"正在输入的那个词"的起始列（1-based 字符列）。
        ///   词还没开始（`New |`）时 <see cref="WordBefore"/> 会给回 caret 本身。
        /// </summary>
        public static bool AfterTypeKeyword(string line, int wordStartCol)
        {
            if (string.IsNullOrEmpty(line)) return false;
            int start = Math.Max(0, Math.Min(wordStartCol - 1, line.Length));
            string before = line.Substring(0, start).TrimEnd();
            if (before.Length == 0) return false;
            // 关键字左边不能再连着标识符字符：`Renew` / `xNew` 不算关键字
            return Regex.IsMatch(before, @"(?<![\p{L}\p{N}_])(?:As|New)$",
                RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// 把模块全文里"正在输入的那个词"挖成空格。
        ///
        /// 为什么要这一步：作用域分析看到的是**当前**文本，而用户打到一半的
        ///   名字（`Dim glo` / `Public g`）本身会被解析成一个声明。于是候选池里
        ///   出现了"这个词自己"，弹出来就是自我提示。挖掉之后池子里只剩
        ///   "已经写完的名字"，自我提示从根上没了。
        ///
        /// 先核对"模块全文那一行那一列确实就是这个词"，对不上就原样返回：
        ///   模块全文有 1.5 秒缓存，偶尔会比光标慢一拍；那时宁可这一次不挖，
        ///   也绝不能拿错位的位置去把别人挖掉（挖坏 `Dim` 会让整个模块的过程
        ///   解析失败）。
        /// </summary>
        public static string BlankToken(string moduleText, int line1, int startCol1, string token)
        {
            if (string.IsNullOrEmpty(moduleText) || string.IsNullOrEmpty(token) ||
                line1 < 1 || startCol1 < 1)
                return moduleText;
            string unified = moduleText.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = unified.Split('\n');
            if (line1 > lines.Length) return moduleText;
            string ln = lines[line1 - 1];
            int s = startCol1 - 1;
            int n = token.Length;
            if (s < 0 || n <= 0 || s + n > ln.Length) return moduleText;
            if (string.Compare(ln, s, token, 0, n, StringComparison.OrdinalIgnoreCase) != 0)
                return moduleText;
            lines[line1 - 1] = ln.Substring(0, s) + new string(' ', n) + ln.Substring(s + n);
            return string.Join("\n", lines);
        }

        // 点号左边允许出现的"能取成员的东西"的末尾字符：`Range("A1")` / `arr(0)` /
        // `x(1)` 收尾是 `)`；`a.b` 收尾是标识符字符。
        private const string MemberOwnerTail = "_)]}";

        // With 块的成员访问：点顶在行首（`    .Size`），或紧跟语句 / 运算符边界
        // （`a = 1: .Size`、`Set f = .Font`、`x = .Left + .Width`）。
        // 这些位置 VBE 同样会弹成员列表。
        private const string WithOwnerBoundary = ":,=(&+-*/\\^><";

        /// <summary>
        /// 光标若停在**成员访问位置**，返回那个点号的 0-based 下标；否则 -1。
        ///
        /// 这是"要不要把这一处让给 VBE"的**唯一判据**（口径："输入 student.，
        /// 点号后面就不要出提示框了"）。只认"光标紧跟点号"是不够的：
        /// `Range("A1").|` / `student.na|` 这类位置会漏判 —— 漏判的后果是
        /// 我们的裸名池照弹、压在 VBE 的成员列表上面。
        ///
        /// 认的位置：
        ///   `student.|` / `student.na|`      标识符.成员
        ///   `Range("A1").|` / `arr(0).x|`    点号左边是 `)` / `]` / `}`
        ///   `    .Size|` / `Set f = .Font|`  With 块的成员访问
        /// 不认的位置（VBE 在这些地方不弹，我们照旧提示）：
        ///   `1.5`（小数点）、`x = "abc".c`（点号贴着一个完整串字面量）、
        ///   注释 / 字符串里的点号。
        /// </summary>
        public static int MemberDotCol(string line, int caretCol)
        {
            if (string.IsNullOrEmpty(line) || caretCol < 1) return -1;
            int i = Math.Min(caretCol - 1, line.Length);

            // 1) 跳过光标左边已经打的成员名。光标紧贴点号时（`student.|`），
            //    那个点号本身不能被当成"正在输入的名字"吃掉。
            int k;
            if (i > 0 && line[i - 1] == '.') k = i - 1;
            else
            {
                int j = i - 1;
                while (j >= 0 && IsIdentChar(line[j])) j--;
                // 2) 成员名左边应当是点号（`X .` 这种带空格的写法也认）
                k = j;
                while (k >= 0 && (line[k] == ' ' || line[k] == '\t')) k--;
                if (k < 0 || line[k] != '.') return -1;
            }

            // 3) 点号左边得是能取成员的东西
            int m = k - 1;
            while (m >= 0 && (line[m] == ' ' || line[m] == '\t')) m--;
            if (m < 0 || WithOwnerBoundary.IndexOf(line[m]) >= 0)
            {
                // With 块 / 点顶行首。注释、字符串里 VBE 不弹，那种点号不算。
                return InCommentOrString(line, k + 1) ? -1 : k;
            }
            if (!(IsIdentChar(line[m]) || MemberOwnerTail.IndexOf(line[m]) >= 0)) return -1;

            // 3b) 点号紧贴一个**收尾引号**：`x = "abc".c` 里的点不是成员访问
            //     （`"abc"` 是完整的串字面量）。只看紧邻那一格，不能看"跳过空白
            //     之后的 m"——`Range("A1").` 的 `)` 在串里，那样写会全误杀。
            if (k - 1 >= 0 && line[k - 1] == '"' && InCommentOrString(line, k)) return -1;

            // 4) `1.5` 这种小数点：点号左边整段是纯数字，不是成员访问
            int s = m;
            while (s >= 0 && IsIdentChar(line[s])) s--;
            string ownerSeg = line.Substring(s + 1, m - s);
            bool allDigits = ownerSeg.Length > 0;
            foreach (char ch in ownerSeg)
                if (!char.IsDigit(ch)) { allDigits = false; break; }
            if (allDigits) return -1;

            // 5) 注释 / 字符串里 VBE 不弹，我们照旧提示
            if (InCommentOrString(line, k + 1)) return -1;
            return k;
        }

        /// <summary>光标是不是停在成员访问位置（`student.|` / `Range("A1").na|`）。</summary>
        public static bool InMemberPosition(string line, int caretCol)
            => MemberDotCol(line, caretCol) >= 0;

        /// <summary>
        /// 光标是不是在 `对象.|` 或 `对象.前|` 这种位置。成立时 root = 点号前那个
        /// 标识符，dotCol = 点号的 1-based 列。
        /// 只认"点号前是一个简单标识符"（`Me` / `c` / `m_arr`）—— `Range("A1").`
        /// 这种要真做表达式求值，不在这一层猜。
        /// </summary>
        public static bool MemberRoot(string line, int caretCol, out string root, out int dotCol)
        {
            root = null; dotCol = -1;
            if (string.IsNullOrEmpty(line)) return false;
            int end = Math.Max(0, Math.Min(caretCol - 1, line.Length));
            int i = end - 1;
            while (i >= 0 && IsIdentChar(line[i])) i--;      // 跳过点号后面已经打的词
            if (i < 0 || line[i] != '.') return false;
            dotCol = i + 1;                                   // 1-based
            int j = i - 1;
            while (j >= 0 && char.IsWhiteSpace(line[j])) j--;
            int ke = j;
            while (j >= 0 && IsIdentChar(line[j])) j--;
            if (ke < 0 || ke - j <= 0) return false;          // 点号前没有标识符
            root = line.Substring(j + 1, ke - j);
            return root.Length > 0;
        }

        /// <summary>
        /// 光标是不是在**注释**里（`'...` 或语句开头的 `Rem ...`）。
        /// 与 <see cref="InCommentOrString"/> 的差别只在"字符串算不算"这一条 ——
        /// 自动配对必须用这一版，原因见 `Editing.CloseOf` 那一段：
        /// 打 `s = "abc` 再按 `"` 时，光标在未闭合的字符串**内部**，那一下正是
        /// 收尾的引号；若按"在字符串里就不配对"提前放弃，VBE 会老实插一个 `"`，
        /// 结果是 `s = "abc""`。正确行为是只补一个收尾引号（判据只看 in_comment）。
        /// </summary>
        public static bool InCommentOnly(string line, int caretCol) =>
            ScanDead(line, caretCol, strCounts: false, remCounts: true);

        /// <summary>光标是不是在注释里（`'` 之后）或字符串里（含未闭合字符串）。</summary>
        public static bool InCommentOrString(string line, int caretCol) =>
            ScanDead(line, caretCol, strCounts: true, remCounts: false);

        /// <summary>
        /// 光标是不是落在"写什么都无所谓"的位置 —— 注释（`'` 或 `Rem`）或字符串里。
        ///
        /// 候选弹窗原来只过 <see cref="InCommentOrString"/>，它不认 `Rem`，
        ///   于是 `Rem 说明 |` 里照样弹全工程候选、按 Tab 还能把注释改写掉。
        ///   两个判据各自漏的那块正好被对方补上，取并集才是全的。
        /// </summary>
        public static bool InDeadZone(string line, int caretCol) =>
            ScanDead(line, caretCol, strCounts: true, remCounts: true);

        /// <summary>
        /// 上面三个判据的公共扫描（三处共用一份实现，改一处就全改到）。
        /// </summary>
        /// <param name="strCounts">未闭合的字符串算不算"死区"。
        ///   true = 算（弹候选没意义）；false = 不算（自动配对要能进到字符串里收尾）。</param>
        /// <param name="remCounts">语句开头的 `Rem` 算不算注释。</param>
        private static bool ScanDead(string line, int caretCol, bool strCounts, bool remCounts)
        {
            if (string.IsNullOrEmpty(line)) return false;
            int end = Math.Max(0, Math.Min(caretCol - 1, line.Length));
            bool inStr = false;
            bool stmtStart = true;          // 还没写出本语句的第一个实义字符
            for (int i = 0; i < end; i++)
            {
                char c = line[i];
                if (inStr)
                {
                    if (c == '"')
                    {
                        if (i + 1 < end && line[i + 1] == '"') { i++; continue; }  // `""` 是转义
                        inStr = false;
                    }
                    continue;
                }
                if (c == '"') { inStr = true; stmtStart = false; continue; }
                if (c == '\'') return true;                       // 注释到行尾，后面不可能再有代码
                if (c == ':') { stmtStart = true; continue; }      // 语句分隔符
                if (c == ' ' || c == '\t') continue;               // 空白不改 stmtStart
                if (!remCounts || !stmtStart) { stmtStart = false; continue; }
                // `Rem` 只有作为语句的第一个词才是注释（`Remember` 不是）
                if (i + 3 <= line.Length &&
                    string.Compare(line, i, "Rem", 0, 3, StringComparison.OrdinalIgnoreCase) == 0 &&
                    (i + 3 == line.Length || line[i + 3] == ' ' || line[i + 3] == '\t'))
                    return true;
                stmtStart = false;
            }
            return strCounts && inStr;
        }

        /// <summary>
        /// 这个词的位置允不允许补一对括号。
        /// 只管"位置"，不管"是什么名字"（那是引擎的活）。
        /// </summary>
        public static bool IsCallSite(string line, int wordStartCol)
        {
            if (line == null) return false;
            int start = Math.Max(0, Math.Min(wordStartCol - 1, line.Length));
            string before = line.Substring(0, start).TrimEnd();
            // 行首（前面只有缩进）同样是调用点：`MsgBox "hi"` / `Split(s)` 都是
            //   合法的语句开头。这里返回 false 的话，"在一行的开头打函数名、
            //   按 Tab”永远补不出括号。
            if (before.Length == 0) return true;
            if (before.EndsWith(".", StringComparison.Ordinal)) return false;
            if (Regex.IsMatch(before, @"\b(?:Sub|Function|Property|Declare|Enum|Type|Event|As|New|On|GoTo|GoSub|Resume|Call)\s*$",
                    RegexOptions.IgnoreCase))
                return false;
            return true;
        }

        // ------------------------------------------------------------------
        // 匹配
        //
        // 模糊匹配的实现【只有一份】，在 Fuzzy.cs（逐字符命中 + 5 档质量）。
        //   再叠一套"前缀 + 词边界跳步"的简版，两个一起用会出现
        //   "排序按 A 算法、高亮按 B 算法"的自相矛盾。
        // ------------------------------------------------------------------

        /// <summary>
        /// "打全了没有"：精确命中【并且】没有更长的候选，才算打全。
        /// 打全就该收起列表 —— 否则按 Tab 会把已经打完的名字再补一遍。
        ///
        /// 单字符例外：一个字母的变量名（`Dim w As Worksheet`、`Dim i&amp;`）
        ///   打完第一个字符就等于"打全"，照这条收的话永远看不到候选 ——
        ///   用户敲 `w` 什么都不弹，敲 `i` 却弹（工程里还有 `itemsArr` 这类
        ///   更长的候选替它撑着），看起来就像"有的变量能提示、有的不能"。
        ///   单字符名字本来也没有"再补一遍"的余地（补上去还是同一个字符），
        ///   所以一个字时不判打全，列表照常给。
        /// </summary>
        public static bool PrefixIsResolved(string word, IList<string> matches)
        {
            if (string.IsNullOrEmpty(word) || matches == null || matches.Count == 0) return false;
            if (word.Length < 2) return false;      // 单字符不收起，见上面的说明
            bool exact = false;
            foreach (var m in matches)
            {
                if (string.Equals(m, word, StringComparison.OrdinalIgnoreCase)) { exact = true; continue; }
                if (m.Length > word.Length &&
                    m.StartsWith(word, StringComparison.OrdinalIgnoreCase))
                    return false;    // 还有更长的候选，没打全
            }
            return exact;
        }

        /// <summary>
        /// 一次文本变化是不是"在光标处插进了一个字符"（用来判断用户真敲了字）。
        /// 命中时 inserted 是插入的那个字符；删除/替换返回 false。
        /// </summary>
        public static bool LineEditInserted(string oldLine, string newLine, out char inserted)
        {
            inserted = '\0';
            oldLine = oldLine ?? "";
            newLine = newLine ?? "";
            if (newLine.Length != oldLine.Length + 1) return false;
            int i = 0;
            while (i < oldLine.Length && oldLine[i] == newLine[i]) i++;
            // 剩下的部分必须逐字相同（中间插了一个字符）
            if (i + 1 > newLine.Length) return false;
            if (!string.Equals(oldLine.Substring(i), newLine.Substring(i + 1), StringComparison.Ordinal))
                return false;
            inserted = newLine[i];
            return true;
        }

        /// <summary>把行尾的空白去掉（比较"有没有实质变化"时用）。</summary>
        public static string Rstrip(string s) => (s ?? "").TrimEnd(' ', '\t', '\r', '\n');

        /// <summary>
        /// 自动弹出候选窗的总判据。四个条件同时成立才弹：
        ///   1. <paramref name="changed"/> —— 快照真的变了；
        ///   2. <paramref name="modSwitched"/> 为假 —— 只是切了代码窗，不是打字；
        ///   3. <paramref name="keyTriggered"/> —— 最近按过键（鼠标右键粘贴、拖拽、
        ///      别的程序用 COM 改模块都不算）；
        ///   4. <paramref name="lineJumped"/> 为假 —— **光标没有跳到别的行**。
        ///
        /// 第 4 条针对的是："有时候鼠标单击到某个变量名，它会
        ///   自动弹提示词框。" 鼠标单击另一行上的变量时，快照的"当前行 + 行文本"
        ///   全变了，看起来和"刚敲了一个字"一模一样；而"最近 1.2 秒按过键"这个
        ///   判据又拦不住它（用户往往是打完字顺手点一下）。真打字时行号绝不会变
        ///   （回车换到的是空行，那种位置本来也没词可提示），方向键上下、Ctrl+Home
        ///   这类"只移动光标"的操作则一定会变 —— 所以行号一变就当"光标移动"处理：
        ///   **只收不弹**。
        /// </summary>
        public static bool AutoTriggerWanted(bool changed, bool modSwitched,
            bool keyTriggered, bool lineJumped)
            => changed && !modSwitched && keyTriggered && !lineJumped;
    }
}
