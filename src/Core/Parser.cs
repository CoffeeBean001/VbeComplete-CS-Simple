// VBA 源码解析：从模块全文里收"工程自己定义的名字"。
//
// 这一层【完全不碰 COM】，只吃一段字符串 —— 于是它能被单独测试，
//   也是"候选池里到底有什么"的唯一事实来源。
//
// 收什么口径）：**只有自己写的名字** —— 过程名（Sub/Function/
//     Property）、变量与常量、参数名。类型名 / 枚举名 / 模块名 / 窗体名 /
//     类名 / VBA 关键字 / 内建函数都不再作为候选（它们在 Scope 里用得上，
//     但那是"作用域判定"和"语法挡板"的事，不进候选池）。
// 不收：注释里和字符串里的词（那是在说人话，不是代码）。

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal sealed class ProcInfo
    {
        public string Name;
        public string Kind;        // sub / function / property
        public int StartLine;      // 1-based，等于 VBE 的行号
        public int EndLine;
        public string ScopeKey;    // 同名过程时带 #序号

        public override string ToString() => $"{Kind} {Name} [{StartLine}-{EndLine}]";
    }

    internal static class Parser
    {
        // 标识符：Unicode 字母/下划线开头，后面可以跟字母数字下划线，可选类型后缀。
        // 必须用 \p{L} 而不是 [A-Za-z]：VBA 允许中文变量名/过程名
        //   （`Sub 汇总()` / `学生 = Split(...)`），旧正则把它们整段跳过，
        //   于是"中文名永远不提示"。
        private const string IDENT = @"[\p{L}_][\p{L}\p{N}_]*[\%\&\#\@\$]?";

        private static readonly Regex ReSub = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Global)[ \t]+)?(?:Static[ \t]+)?Sub[ \t]+(" + IDENT + @")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReFunc = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Global)[ \t]+)?(?:Static[ \t]+)?Function[ \t]+(" + IDENT + @")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReProp = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Global)[ \t]+)?(?:Static[ \t]+)?Property[ \t]+(?:Get|Let|Set)[ \t]+(" + IDENT + @")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReEndProc = new Regex(
            @"^[ \t]*End[ \t]+(?:Sub|Function|Property)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// 单行过程：过程头之后**同一行里**还有 `: End Sub`（/Function/Property）。
        /// VBE 允许 `Sub A(): x = 1: End Sub`，而 <see cref="ReEndProc"/> 锚在行首，
        /// 认不出行尾那个 `End`。
        /// </summary>
        private static readonly Regex ReInlineEnd = new Regex(
            @":[ \t]*End[ \t]+(?:Sub|Function|Property)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReOptionExplicit = new Regex(
            @"^[ \t]*Option[ \t]+Explicit\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>按物理行切分（VBE 的行号就是物理行号，不能把续行合并掉）。</summary>
        public static string[] SplitLines(string text)
        {
            if (text == null) return Array.Empty<string>();
            string t = text.Replace("\r\n", "\n").Replace('\r', '\n');
            // 末尾的换行会切出一个空元素，VBE 的 CountOfLines 也算这一行，保留
            return t.Split('\n');
        }

        /// <summary>去掉注释与字符串内容，只留"代码形状"的部分。</summary>
        public static string CodeOnly(string line)
        {
            if (string.IsNullOrEmpty(line)) return "";
            var sb = new StringBuilder(line.Length);
            bool inStr = false;
            bool stmtStart = true;              // 行首算语句起点
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inStr)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { i++; continue; }
                        inStr = false;
                    }
                    sb.Append(' ');
                    continue;
                }
                if (c == '"') { inStr = true; sb.Append(' '); continue; }
                if (c == '\'') break;                      // 行内注释

                // `Rem` 出现在**语句起点**（行首，或者冒号之后）就是注释，
                //   不再只认行首。`x = 1: Rem 临时占位` 里的 `临时占位` 否则会被
                //   当成代码，进而作为"未声明变量"进候选池。
                //   长度保持守恒（余下全填空格）：调用方虽然不依赖，但少一条隐患。
                if (stmtStart && IsRemAt(line, i))
                {
                    sb.Append(' ', line.Length - i);
                    break;
                }

                if (c == ':') { stmtStart = true; sb.Append(c); continue; }
                if (c != ' ' && c != '\t') stmtStart = false;
                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>line[i] 起是不是一个独立的 `Rem` 关键字（后面不是字母数字）。</summary>
        private static bool IsRemAt(string line, int i)
        {
            if (i + 2 >= line.Length) return false;
            if ((line[i] != 'R' && line[i] != 'r') ||
                (line[i + 1] != 'E' && line[i + 1] != 'e') ||
                (line[i + 2] != 'M' && line[i + 2] != 'm')) return false;
            int after = i + 3;
            return after >= line.Length || !char.IsLetterOrDigit(line[after]);
        }

        /// <summary>模块里所有过程（含同名去重用的 ScopeKey）。</summary>
        public static List<ProcInfo> ExtractProcedures(string text)
        {
            var list = new List<ProcInfo>();
            var lines = SplitLines(text);
            var open = new List<ProcInfo>();
            var used = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < lines.Length; i++)
            {
                string code = CodeOnly(lines[i]);
                if (code.Length == 0) continue;

                Match m = ReSub.Match(code);
                string kind = "sub";
                if (!m.Success) { m = ReFunc.Match(code); kind = "function"; }
                if (!m.Success) { m = ReProp.Match(code); kind = "property"; }

                if (m.Success)
                {
                    string name = m.Groups[1].Value.TrimEnd('%', '&', '!', '#', '@', '$');
                    int n;
                    used.TryGetValue(name, out n);
                    used[name] = ++n;

                    // 新过程头出现时，上一个**还没等到 `End Sub`** 的过程要在这里
                    //   强制收尾。EndLine 若一直停留在"模块最后一行"，后面
                    //   整段代码都算在它名下（模块级声明被当成过程局部、别的过程的变量
                    //   跨过程泄漏、`Ctrl+↓` 也跳错位置）。
                    if (open.Count > 0)
                    {
                        foreach (var prev in open)
                            if (prev.EndLine > i) prev.EndLine = Math.Max(prev.StartLine, i);
                        open.Clear();
                    }

                    var p = new ProcInfo
                    {
                        Name = name,
                        Kind = kind,
                        StartLine = i + 1,
                        EndLine = lines.Length,
                        ScopeKey = n > 1 ? name + "#" + n : name,
                    };
                    list.Add(p);

                    // 单行过程 `Sub A(): x = 1: End Sub` —— 行尾的 End 不是
                    //   行首，`ReEndProc`（锚在 `^`）认不出来，于是这个过程一路吃到
                    //   模块末尾。这里按"同一行里还有 `: End Sub`"判掉，别进 open。
                    if (ReInlineEnd.IsMatch(code))
                    {
                        p.EndLine = i + 1;
                        continue;
                    }

                    open.Add(p);
                    continue;
                }

                if (ReEndProc.IsMatch(code) && open.Count > 0)
                {
                    open[open.Count - 1].EndLine = i + 1;
                    open.RemoveAt(open.Count - 1);
                }
            }
            return list;
        }

        /// <summary>第 line 行（1-based）落在哪个过程里；不在任何过程里返回 null。</summary>
        public static ProcInfo ProcOfLine(List<ProcInfo> procs, int line)
        {
            if (procs == null) return null;
            foreach (var p in procs)
                if (line >= p.StartLine && line <= p.EndLine) return p;
            return null;
        }

        /// <summary>按顶层逗号切分（不切括号里面的逗号）。</summary>
        public static List<string> SplitTopLevel(string s)
        {
            var outp = new List<string>();
            if (s == null) return outp;
            int depth = 0;
            var cur = new StringBuilder();
            foreach (char c in s)
            {
                if (c == '(') depth++;
                else if (c == ')') depth = Math.Max(0, depth - 1);
                if (c == ',' && depth == 0) { outp.Add(cur.ToString()); cur.Clear(); continue; }
                cur.Append(c);
            }
            outp.Add(cur.ToString());
            return outp;
        }

        public static bool OptionExplicit(string text)
        {
            if (text == null) return false;
            foreach (var raw in SplitLines(text))
                if (ReOptionExplicit.IsMatch(raw)) return true;
            return false;
        }
    }
}
