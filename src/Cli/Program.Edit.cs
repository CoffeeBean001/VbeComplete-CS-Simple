// 编辑手感相关的离线诊断：--writetest、--paretest、--carettest。
//
// 这些命令只在命令行里跑，不进正常启动路径。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    internal static partial class Program
    {
        /// <summary>
        /// 写入链路验证：在模块【末尾】追加一行测试代码 -> 用补全改它 -> 删掉。
        /// 全程可逆，最后逐字节比对原文；比对失败会把原文整段写进日志。
        /// 只由用户显式运行 `--writetest` 触发，正常运行一次都不会动用户代码。
        /// </summary>
        private static int WriteTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool ok = true;

            sb.AppendLine("=== write test ===");
            sb.AppendLine("module lines = " + count + "  text len = " + (orig == null ? -1 : orig.Length));
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "writetest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            try
            {
                const string probe = "    Dim abc";
                if (!be.InsertLines(count + 1, probe))
                {
                    sb.AppendLine("InsertLines FAILED");
                    ok = false;
                }
                else
                {
                    string back = be.LineAt(count + 1);
                    sb.AppendLine("insert readback = [" + back + "]");
                    bool same = back == probe;
                    if (!same) ok = false;
                    sb.AppendLine("insert roundtrip = " + same);

                    bool w = be.ApplyCompletion(count + 1, 9, 12, "abcd", false, false, null);
                    string after = be.LineAt(count + 1);
                    sb.AppendLine("apply ok=" + w + "  result=[" + after + "]");
                    sb.AppendLine("apply result correct = " + (after == "    Dim abcd"));
                    if (after != "    Dim abcd") ok = false;

                    be.SetCaret(count + 1, 13);
                    var s2 = be.Snapshot();
                    string got = s2 == null ? "null" : s2.Line + ":" + s2.CaretCol;
                    sb.AppendLine("caret after SetCaret(13) = " + got + " (want 13)");
                    if (s2 == null || s2.CaretCol != 13) ok = false;
                }
            }
            finally
            {
                // 无论成败都要把测试行删干净（最多重试 3 次）
                for (int i = 0; i < 3; i++)
                {
                    int c2 = be.LineCount();
                    if (c2 <= count) break;
                    be.DeleteLines(count + 1, c2 - count);
                }
            }

            string now = be.ModuleText();
            bool restored = now == orig;
            sb.AppendLine("restored = " + restored);
            if (!restored)
            {
                ok = false;
                sb.AppendLine("!!! 未能逐字节还原，原文如下（请手工检查模块末尾） !!!");
                sb.AppendLine(orig);
                sb.AppendLine("!!! 现在的全文 !!!");
                sb.AppendLine(now);
            }
            sb.AppendLine(ok && restored ? "RESULT: PASS" : "RESULT: FAIL");

            File.WriteAllText(Path.Combine(Settings.Home, "writetest.log"), sb.ToString(), Encoding.UTF8);
            return ok && restored ? 0 : 1;
        }

        /// <summary>
        /// 括号补全验证（真机）：走 Engine.AcceptName 这条**真路径**（不是直接调
        /// Backend.ApplyCompletion），确认"选中自己写的 Function"真的会得到一对括号、
        /// 光标真的落在两个括号中间；Sub 补空格，变量什么都不补。
        ///
        /// <summary>日志用：安全截一段字符串（越界不抛，换行写成 \n）。</summary>
        private static string SafeSlice(string s, int from, int len)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (from < 0) from = 0;
            if (from >= s.Length) return "";
            string part = s.Substring(from, Math.Min(len, s.Length - from));
            return part.Replace("\r", "\\r").Replace("\n", "\\n");
        }

        /// 为什么要单开一条：验"选中的是 vba 函数提示词，没有自动补全双括号"这一档。
        ///   括号逻辑分散在四处（Followup 判档位 → IsCallSite 判位置 →
        ///   InCallStatement 判 Call 语句 → ApplyCompletion 落光标），只有真机
        ///   走一遍才知道断在哪一环。全程只动模块末尾的临时行，最后逐字节还原。
        /// 候选池里已经没有内建词了，所以档位用**注入的一段合成源码**
        ///   PrimePool 出来（`我的函数` = Function / `我的过程` = Sub / `我的值` = 变量），
        ///   不再依赖词库文件是否在。
        /// </summary>
        private static int ParenTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool ok = true;

            sb.AppendLine("=== paren test (真机) ===");
            sb.AppendLine("module lines = " + count);
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "paretest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            var eng = new Engine { Backend = be };
            // 档位靠注入的合成源码，不再依赖词库文件：Function -> 补括号、
            // Sub -> 补空格、Property / 变量 -> 都不补。
            eng.PrimePool(
                "Public Function myFunc()\r\nEnd Function\r\n" +
                "Public Sub myProc()\r\nEnd Sub\r\n" +
                "Public Property Get myProp() As Long\r\nEnd Property\r\n",
                "Module1");

            // 档位判据先单独对一遍（不碰 COM）
            bool wp, ws;
            eng.Followup("myFunc", null, out wp, out ws);
            sb.AppendLine("Followup(myFunc 函数) -> parens=" + wp + " space=" + ws);
            ok &= wp && !ws;
            eng.Followup("myProc", null, out wp, out ws);
            sb.AppendLine("Followup(myProc 过程) -> parens=" + wp + " space=" + ws);
            ok &= !wp && ws;
            eng.Followup("myProp", null, out wp, out ws);
            sb.AppendLine("Followup(myProp 属性) -> parens=" + wp + " space=" + ws);
            ok &= !wp && !ws;
            eng.Followup("someVar", null, out wp, out ws);
            sb.AppendLine("Followup(someVar 变量) -> parens=" + wp + " space=" + ws);
            ok &= !wp && !ws;

            // (探针行, 光标列, 候选名, 期望行, 期望光标列)
            var cases = new[]
            {
                // 一行的开头（只有缩进）
                ("    myFu", 9, "myFunc", "    myFunc()", 12),
                // 赋值右侧
                ("    x = myFu", 13, "myFunc", "    x = myFunc()", 16),
                // Sub：补空格，不补括号
                ("    myPr", 9, "myProc", "    myProc ", 12),
                // Property：什么都不补
                ("    myPr", 9, "myProp", "    myProp", 11),
                // （对标 IDEA）：光标停在词**中间**时，替换的是整个标识符
                //   —— 只换光标左边那一截会得到 `myFuncFunc`。
                ("    myFunx", 9, "myFunc", "    myFunc()", 12),
            };

            try
            {
                foreach (var c in cases)
                {
                    be.InsertLines(count + 1, c.Item1);
                    be.SetCaret(count + 1, c.Item2);
                    bool acc = eng.AcceptName(c.Item3);
                    string line = be.LineAt(count + 1) ?? "";
                    var snap = be.Snapshot();
                    int caret = snap == null ? -1 : snap.CaretCol;

                    bool lineOk = line == c.Item4;
                    bool caretOk = caret == c.Item5;
                    if (!acc || !lineOk || !caretOk) ok = false;
                    sb.AppendLine("  [" + (acc && lineOk && caretOk ? "PASS" : "FAIL") + "] "
                        + c.Item1.TrimEnd() + " + Tab " + c.Item3);
                    sb.AppendLine("         行 = [" + line + "]  want [" + c.Item4 + "]"
                        + (lineOk ? "" : "  <- 行不对"));
                    sb.AppendLine("         光标 = " + caret + "  want " + c.Item5
                        + (caretOk ? "" : "  <- 落点不对"));

                    // 立刻删掉这一行，下一个用例从同一个基准行数开始
                    for (int i = 0; i < 3; i++)
                    {
                        int c2 = be.LineCount();
                        if (c2 <= count) break;
                        be.DeleteLines(count + 1, c2 - count);
                    }
                }
            }
            finally
            {
                for (int i = 0; i < 3; i++)
                {
                    int c2 = be.LineCount();
                    if (c2 <= count) break;
                    be.DeleteLines(count + 1, c2 - count);
                }
            }

            string now = be.ModuleText();
            bool restored = now == orig;
            sb.AppendLine("restored = " + restored);
            if (!restored)
            {
                ok = false;
                int n = Math.Min(now == null ? 0 : now.Length, orig == null ? 0 : orig.Length);
                int at = 0;
                while (at < n && now[at] == orig[at]) at++;
                sb.AppendLine("首个差异位置 = " + at + "（len now=" + (now == null ? -1 : now.Length)
                    + " orig=" + (orig == null ? -1 : orig.Length) + "）");
                int from = Math.Max(0, at - 40);
                sb.AppendLine("orig[" + from + "..] = [" + SafeSlice(orig, from, 120) + "]");
                sb.AppendLine("now [" + from + "..] = [" + SafeSlice(now, from, 120) + "]");
            }
            sb.AppendLine(ok && restored ? "RESULT: PASS" : "RESULT: FAIL");

            File.WriteAllText(Path.Combine(Settings.Home, "paretest.log"), sb.ToString(), Encoding.UTF8);
            return ok && restored ? 0 : 1;
        }

        /// <summary>
        /// 光标落点验证（真机）：模拟"打到一半 + Tab 补全"，读回 VBE 真实光标列，
        /// 跟"按读回来的真行算出来的列"对数，并把光标左右两截文本打出来。
        /// 只在 `--carettest` 时执行；动的是模块末尾的临时行，finally 里删干净
        /// 并逐字节比对原文。
        /// </summary>
        private static int CaretTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool allOk = true;

            sb.AppendLine("=== caret test ===");
            sb.AppendLine("module lines = " + count + "  text len = " + (orig == null ? -1 : orig.Length));
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "carettest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            // (标签, 插入的行, 已输入的词, 候选)
            // ⚠️ InsertLines 自己也会触发 VBE 的行格式化（`student=spl` 读回是
            //   `student = spl`），所以列**必须按读回来的真行**算 —— 这正是
            //   真机与离线用例的差别，也是这个 bug 之前没被发现的原因。
            var cases = new[]
            {
                ("顶格无缩进",        "student=spl",  "spl", "Split"),
                ("4 空格缩进",        "    student=spl", "spl", "Split"),
                ("Tab 缩进",          "\tstudent=spl", "spl", "Split"),
                ("Tab 缩进 + 全角名",  "\t学生=spl",   "spl", "Split"),
                ("右边还有同名子串",   "    abc=spl+Splitter", "spl", "Split"),
            };

            try
            {
                foreach (var c in cases)
                {
                    int line = count + 1;
                    if (!be.InsertLines(line, c.Item2))
                    {
                        sb.AppendLine(c.Item1 + ": InsertLines FAILED");
                        allOk = false;
                        continue;
                    }
                    string inserted = be.LineAt(line) ?? "";
                    int at0 = WholeWordIndex(inserted, c.Item3);
                    if (at0 < 0)
                    {
                        sb.AppendLine(c.Item1 + ": 读回行里找不到已输入的词 [" + inserted + "]");
                        allOk = false;
                        be.DeleteLines(line, 1);
                        continue;
                    }
                    int wordStart = at0 + 1;                  // 1-based 字符列
                    int endCol = at0 + 1 + c.Item3.Length;

                    // 这一条验的是【Engine 真正用的那对换算】——
                    //   VBE 的 `CaretCol` 是显示列，直接拿去切字符下标，含全角/Tab 的行
                    //   会切出空词，症状是"这一行根本不弹候选"。所以用"打完还没补全"
                    //   的那一行做样本：把词尾的显示列换算成字符列再取词，必须还是原词。
                    int wordEndDisp = Rules.DispCol(inserted, at0 + c.Item3.Length);
                    int ws;
                    string wordSeen = Rules.WordBefore(inserted,
                        Rules.CharColFromDisp(inserted, wordEndDisp), out ws);
                    bool okWord = string.Equals(wordSeen, c.Item3, StringComparison.OrdinalIgnoreCase);

                    bool w = be.ApplyCompletion(line, wordStart, endCol, c.Item4, true, false, null);
                    string actual = be.LineAt(line) ?? "";
                    var snap = be.Snapshot();
                    int got = snap == null ? -1 : snap.CaretCol;

                    // 判据必须【独立】—— 用 DispCol 算期望列，等于拿
                    //   被测的同一套语义去验它自己，于是"落在 `(` 左边"也能全绿。
                    //   现在直接看光标左右两边是什么字符：
                    //   补过括号 ⇒ 左邻必须是 `(`、右邻必须是 `)`（= 夹在括号中间）。
                    int idx = actual.IndexOf('(');
                    bool hasParens = idx >= 0;
                    int want = hasParens
                        ? Rules.DispCol(actual, idx + 1)          // `)` 之前
                        : Rules.DispCol(actual, at0 + c.Item4.Length);

                    // 约定：列 C = "字符 C-1 之前"，
                    // 但 C 是【显示列】—— 全角/ Tab 行上不能直接减 1，要先换回字符列。
                    int at = Rules.CharColFromDisp(actual, got) - 1;
                    string left = (at < 0 || at > actual.Length) ? "?" : actual.Substring(0, at);
                    string right = (at < 0 || at > actual.Length) ? "?" : actual.Substring(at);
                    char chLeft = (at >= 1 && at <= actual.Length) ? actual[at - 1] : '\0';
                    char chRight = (at >= 0 && at < actual.Length) ? actual[at] : '\0';
                    bool betweenParens = hasParens && chLeft == '(' && chRight == ')';

                    bool ok = w && got == want && okWord && betweenParens;
                    if (!ok) allOk = false;

                    sb.AppendLine("--- " + c.Item1 + " ---");
                    sb.AppendLine("  inserted  = [" + inserted.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  actual    = [" + actual.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  apply ok  = " + w + "   (wordStart=" + wordStart + " endCol=" + endCol + ")");
                    sb.AppendLine("  caret col = " + got + "   want = " + want + "   " + (got == want ? "OK" : "MISMATCH"));
                    sb.AppendLine("  parens    = 左[" + Esc(chLeft) + "] 右[" + Esc(chRight) + "]  " +
                                  (betweenParens ? "OK(夹在括号中间)" : (hasParens ? "MISMATCH(不在括号中间)" : "无括号")));
                    sb.AppendLine("  engine取词= [" + wordSeen + "]  " + (okWord ? "OK" : "MISMATCH"));
                    sb.AppendLine("  left      = [" + left.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  right     = [" + right.Replace("\t", "\\t") + "]");

                    be.DeleteLines(line, 1);
                }
            }
            finally
            {
                for (int i = 0; i < 3; i++)
                {
                    int c2 = be.LineCount();
                    if (c2 <= count) break;
                    be.DeleteLines(count + 1, c2 - count);
                }
            }

            string now = be.ModuleText();
            bool restored = now == orig;
            sb.AppendLine("restored = " + restored);
            if (!restored) allOk = false;
            sb.AppendLine(allOk ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "carettest.log"), sb.ToString(), Encoding.UTF8);
            return allOk ? 0 : 1;
        }

        /// <summary>把单个字符渲染成日志里看得见的样子（空格/Tab/空位都要能区分）。</summary>
        private static string Esc(char c) =>
            c == '\0' ? "∅" : c == '\t' ? "\\t" : c == ' ' ? "␣" : c.ToString();

        /// <summary>整词查找（前后都不是标识符字符），找不到返回 -1。测试用。</summary>
        private static int WholeWordIndex(string text, string word)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(word)) return -1;
            int i = 0;
            while (i <= text.Length - word.Length)
            {
                int hit = text.IndexOf(word, i, StringComparison.OrdinalIgnoreCase);
                if (hit < 0) return -1;
                bool okL = hit == 0 || !Rules.IsIdentChar(text[hit - 1]);
                bool okR = hit + word.Length >= text.Length || !Rules.IsIdentChar(text[hit + word.Length]);
                if (okL && okR) return hit;
                i = hit + 1;
            }
            return -1;
        }
    }
}
