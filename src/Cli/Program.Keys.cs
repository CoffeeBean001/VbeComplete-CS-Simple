// 热键命令的真机诊断：--hktest。
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
        /// 真机验证四个命令：Ctrl+Shift+Enter / Ctrl+D / Ctrl+K / Ctrl+/。
        /// （`Ctrl+Shift+↑/↓` 的用例 随功能一起删除；`Ctrl+↑/↓` 由 --enterprobe 的
        ///   用例 I 走真钩子验证。）
        ///
        /// 为什么要这个：验"快捷键全都没了"卡在哪一环。按键有没有走到钩子，在沙箱里
        ///   没法用真键盘试（钩子按设计不理注入键）；但"命令本身能不能干活"可以
        ///   100% 验证 —— 因为这里调的就是钩子调的那几个 Commands.* 函数。
        ///   命令全过 ⇒ 问题只可能在路由（焦点判定 / 修饰键 / 别的软件抢钩子），
        ///   App 里那两行"热键被跳过 / 热键未命中"日志就是给这种情况准备的。
        ///
        /// 全程只动模块末尾临时追加的 6 行，最后删干净并逐字节比对原文。
        /// </summary>
        private static int HkTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int baseCount = be.LineCount();
            bool allOk = true;

            sb.AppendLine("=== hotkey test (真机) ===");
            sb.AppendLine("module lines = " + baseCount + "  len = " + (orig == null ? -1 : orig.Length));
            if (baseCount <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "hktest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            var block = new[]
            {
                "Sub VbcHkTest()",
                "    Dim n As Long",
                "    n = 1",
                "    n = 2",
                "",
                "End Sub",
            };
            for (int i = 0; i < block.Length; i++)
                be.InsertLines(baseCount + 1 + i, block[i]);

            int Ldim = baseCount + 2, L1 = baseCount + 3, L2 = baseCount + 4;
            string dimText = be.LineAt(Ldim) ?? "";
            string oneText = be.LineAt(L1) ?? "";
            string twoText = be.LineAt(L2) ?? "";
            sb.AppendLine("scratch: dim=[" + dimText + "] one=[" + oneText + "] two=[" + twoText + "]");

            Action<string, bool, string> check = (name, ok, detail) =>
            {
                if (!ok) allOk = false;
                sb.AppendLine((ok ? "  PASS  " : "  FAIL  ") + name + "   " + detail);
            };

            try
            {
                // ---- 0. SetCaret 的列语义（五个命令全都踩在它上面）----
                {
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    int c1 = be.Snapshot()?.CaretCol ?? -1;
                    be.SetCaret(L1, 3);
                    be.InvalidateSnapshot();
                    int c3 = be.Snapshot()?.CaretCol ?? -1;
                    check("SetCaret(.,1)/(.,3) 列语义", c1 == 1 && c3 == 3,
                        "读回 col=" + c1 + "/" + c3 + "  want 1/3");
                }

                // ---- 1. Ctrl+D 复制行 ----
                {
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    int before = be.LineCount();
                    bool r = Commands.DuplicateLine(be);
                    int after = be.LineCount();
                    var snap = be.Snapshot();
                    string a = be.LineAt(L1) ?? "", b = be.LineAt(L1 + 1) ?? "";
                    bool ok = r && after == before + 1
                              && Rules.Rstrip(a) == Rules.Rstrip(b)
                              && Rules.Rstrip(b).Trim() == "n = 1"
                              && snap != null && snap.Line == L1 + 1;
                    check("Ctrl+D 复制行", ok,
                        "ret=" + r + " lines " + before + "->" + after +
                        " dup=[" + b + "] caretLine=" + (snap?.Line ?? -1));
                    // 复原：副本还在 L1+1
                }

                // ---- 2. Ctrl+K 删行（光标落到上一行行尾）----
                {
                    int before = be.LineCount();
                    be.SetCaret(L1 + 1, 1);           // 刚复制出来的那一行
                    be.InvalidateSnapshot();
                    bool r = Commands.DeleteLine(be);
                    int after = be.LineCount();
                    var snap = be.Snapshot();
                    string now = be.LineAt(L1 + 1) ?? "";
                    bool ok = r && after == before - 1
                              && Rules.Rstrip(now) == Rules.Rstrip(twoText)
                              && snap != null && snap.Line == L1;
                    check("Ctrl+K 删行落到上一行", ok,
                        "ret=" + r + " lines " + before + "->" + after +
                        " next=[" + now + "] caretLine=" + (snap?.Line ?? -1) + " want=" + L1);
                }

                // ---- 3. Ctrl+/ 注释开关 ----
                {
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    var s0 = be.Snapshot();
                    int sl, sc, el, ec;
                    bool hasSel = be.TryGetSelection(out sl, out sc, out el, out ec);
                    bool r1 = Commands.ToggleComment(be);
                    string c1 = be.LineAt(L1) ?? "";
                    be.InvalidateSnapshot();
                    bool r2 = Commands.ToggleComment(be);
                    string c2 = be.LineAt(L1) ?? "";
                    bool ok = r1 && r2
                              && c1.TrimStart().StartsWith("'", StringComparison.Ordinal)
                              && c1.Contains("n = 1")
                              && Rules.Rstrip(c2) == Rules.Rstrip(oneText);
                    check("Ctrl+/ 注释开关", ok,
                        "ret=" + r1 + "/" + r2 + " 加后=[" + c1 + "] 还原后=[" + c2 + "]");
                    sb.AppendLine("        snapLine=" + (s0?.Line ?? -1) + " sel=" + hasSel +
                                  "(" + sl + "," + sc + "," + el + "," + ec + ")" +
                                  " 原行=[" + oneText + "] 期望还原=[" + Rules.Rstrip(oneText) + "]");
                }

                // ---- 4. Ctrl+Shift+Enter 上方插行 ----
                {
                    int before = be.LineCount();
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    bool r = Commands.InsertLineAbove(be);
                    int after = be.LineCount();
                    string ins = be.LineAt(L1) ?? "";
                    string below = be.LineAt(L1 + 1) ?? "";
                    int indent = Commands.LeadingWs(oneText).Length;
                    bool ok = r && after == before + 1
                              && Rules.Rstrip(below) == Rules.Rstrip(oneText)
                              && Commands.LeadingWs(ins).Length == indent;
                    check("Ctrl+Shift+Enter 上方插行", ok,
                        "ret=" + r + " lines " + before + "->" + after +
                        " ins=[" + ins + "] 缩进=" + Commands.LeadingWs(ins).Length + " want=" + indent +
                        " below=[" + below + "] want=[" + oneText + "]");
                    // 无论成败都要把这行临时行删掉 —— 否则后面几条用例的行号全偏，
                    //   一个失败会伪装成三个失败（第一次跑就是这样）。
                    if (after > before)
                    {
                        be.SetCaret(L1, 1);
                        be.InvalidateSnapshot();
                        Commands.DeleteLine(be);
                    }
                }

                // ==========================================================
                // 以下六条是 新移植的"编辑手感"命令。每一条都自己造一小段
                // 临时代码、验完立刻整段删掉 —— 不依赖上面那 6 行的行号。
                // 原来这里的第 5、6 条验的是 `Ctrl+Shift+↑/↓`（上方最近的
                //   声明行 / 下方第一个空白行）。用户明确说用不到，功能已整组删除，
                //   用例一起删掉 —— 留着会编译不过（函数没了）。
                // ==========================================================
                Func<string[], int> append = lines =>
                {
                    int at = be.LineCount() + 1;
                    // 倒着插（全部插在 at 这一位）：正着插的话，`""`（空行）会在
                    //   某一瞬间成为**模块最后一行**，而 VBE 对"末尾的空行"处理
                    //   并不可靠（可能直接吞掉），后面几条用例的行号就全偏了。
                    for (int i = lines.Length - 1; i >= 0; i--) be.InsertLines(at, lines[i]);
                    return at;
                };
                Action dropFrom = () =>
                {
                    for (int i = 0; i < 3; i++)
                    {
                        int c = be.LineCount();
                        if (c <= baseCount) break;
                        be.DeleteLines(baseCount + 1, c - baseCount);
                    }
                };

                // ---- 7. Shift+Enter：下方插一行，缩进抄本行 ----
                {
                    int at = append(new[] { "Sub VbcNl()", "    n = 1", "End Sub" });
                    be.SetCaret(at + 1, 1);
                    be.InvalidateSnapshot();
                    bool r = be.NewLineBelow(false, false, false);
                    string ins = be.LineAt(at + 2) ?? "";
                    var snap = be.Snapshot();
                    bool ok = r && ins == "    " && Rules.Rstrip(be.LineAt(at + 1) ?? "") == "    n = 1"
                              && snap != null && snap.Line == at + 2;
                    check("Shift+Enter 下方插行", ok,
                        "ret=" + r + " 新行=[" + ins.Replace("\t", "\\t") + "] 期望=[    ] caret=" + (snap?.Line ?? -1));
                    dropFrom();
                }

                // ---- 8. 回车自动缩进 + 补收尾 ----
                {
                    int at = append(new[] { "Sub VbcIn()", "    For i = 1 To 3", "End Sub" });
                    // 行尾 = "字符数 + 1"（SetCaret 的列是"第几个字符之前"）
                    be.SetCaret(at + 1, (be.LineAt(at + 1) ?? "").Length + 1);
                    be.InvalidateSnapshot();
                    bool r = be.NewLineBelow(true, true, true);
                    string line2 = be.LineAt(at + 2) ?? "";
                    string line3 = be.LineAt(at + 3) ?? "";
                    var snap = be.Snapshot();
                    bool ok = r && line2 == "        " && Rules.Rstrip(line3) == "    Next"
                              && snap != null && snap.Line == at + 2;
                    check("回车自动缩进 + 补 Next", ok,
                        "ret=" + r + " 新行=[" + line2 + "] 下一行=[" + line3 + "] caret=" + (snap?.Line ?? -1));
                    dropFrom();
                }

                // ---- 9. 退格删整行空白 ----
                {
                    int at = append(new[] { "Sub VbcBs()", "    ", "    n = 2", "End Sub" });
                    string raw = be.LineAt(at + 1) ?? "";
                    be.SetCaret(at + 1, raw.Length + 1);
                    be.InvalidateSnapshot();
                    int before = be.LineCount();
                    bool r = be.DeleteBlankLineHere();
                    int after = be.LineCount();
                    if (raw.Length == 0)
                    {
                        // VBE 把只有空白的行 trim 成真空行了 —— 那时"原生退格"的结果
                        // 与删整行一样，我们不接管是**正确**的
                        check("退格删整行空白（跳过：VBE 已 trim 成真空行）", !r,
                            "raw 为空 -> 不接管，ret=" + r);
                    }
                    else
                    {
                        check("退格删整行空白", r && after == before - 1 && (be.LineAt(at + 1) ?? "").Trim() == "n = 2",
                            "ret=" + r + " lines " + before + "->" + after + " 下一行=[" + be.LineAt(at + 1) + "]");
                    }
                    dropFrom();
                }

                // ---- 10. 退格删一对空括号 ----
                {
                    int at = append(new[] { "Sub VbcDp()", "    x = f()", "End Sub" });
                    be.SetCaret(at + 1, 11);               // 夹在 `(` 与 `)` 中间
                    be.InvalidateSnapshot();
                    bool r = be.DeletePairHere();
                    string now = be.LineAt(at + 1) ?? "";
                    var snap = be.Snapshot();
                    bool ok = r && Rules.Rstrip(now) == "    x = f" && snap != null && snap.CaretCol == 10;
                    check("退格删一对空括号", ok,
                        "ret=" + r + " 行=[" + now + "] caretCol=" + (snap?.CaretCol ?? -1) + " want=10");
                    dropFrom();
                }

                // ---- 11. 自动配对括号 ----
                {
                    int at = append(new[] { "Sub VbcAu()", "    abc" });
                    be.SetCaret(at + 1, (be.LineAt(at + 1) ?? "").Length + 1);   // 行尾
                    be.InvalidateSnapshot();
                    bool r = be.InsertPair('(');
                    string now = be.LineAt(at + 1) ?? "";
                    var snap = be.Snapshot();
                    int col = snap == null ? -1 : snap.CaretCol;
                    // "    abc()" 共 9 个字符，光标落在 `(` 与 `)` **之间** = 第 9 列
                    // （第 k 列 = "第 k 个字符之前"，`)` 是第 9 个字符）
                    bool ok = r && Rules.Rstrip(now) == "    abc()" && col == 9;
                    check("自动配对括号", ok,
                        "ret=" + r + " 行=[" + now + "] caretCol=" + col + " want=9");
                    dropFrom();
                }

                // ---- 12. 插入代码片段（多行） ----
                {
                    // 中间留一行**空行**当落点：片段的落点模型是"整行拼接"
                    //   （前半行 + 片段首行 / 片段末行 + 后半行），光标停在有内容的
                    //   行尾时会拼成 `    x = 0    a = 1` —— 那是设计如此，不是 bug。
                    //   拿空行当落点才是用户真实用法：点一下空行，插入。
                    int at = append(new[] { "Sub VbcSn()", "", "End Sub" });
                    be.SetCaret(at + 1, 1);
                    be.InvalidateSnapshot();
                    int before = be.LineCount();
                    bool r = be.InsertSnippet("    a = 1\r\n    b = 2");
                    int after = be.LineCount();
                    string l1 = be.LineAt(at + 1) ?? "";
                    string l2 = be.LineAt(at + 2) ?? "";
                    string l3 = be.LineAt(at + 3) ?? "";
                    bool ok = r && after == before + 1
                              && Rules.Rstrip(l1) == "    a = 1"
                              && Rules.Rstrip(l2) == "    b = 2"
                              && Rules.Rstrip(l3) == "End Sub";
                    check("插入多行片段", ok,
                        "ret=" + r + " lines " + before + "->" + after +
                        " [" + l1 + "][" + l2 + "][" + l3 + "]");
                    dropFrom();
                }

                // ---- 13. （原 Ctrl+Shift+L 格式化缩进）----
                // 用户要求删除该功能，用例与实现一起移除。编号留着不复用，
                // 免得跟历史日志 / 提交信息里的"第 13 条"对不上。

                // ---- 14. 光标跳到别的行（鼠标单击 / 方向键）不许弹窗 ----
                // 症状："有时候鼠标单击到某个变量名，它会自动弹提示词框。"
                // 这条验的是那个判据**依赖的事实**：真机上"光标换行"确实会让
                // Snapshot.Line 变掉（判据靠它把鼠标操作和打字分开）。
                {
                    var blk = new[] { "Sub VbcJump()", "    n = 1", "    n = 2", "End Sub" };
                    int at = append(blk);
                    be.SetCaret(at + 1, 1);                       // 模拟"刚才在这一行打字"
                    be.InvalidateSnapshot();
                    var a = be.Snapshot();
                    // 模拟鼠标单击下一行那个变量名的词尾
                    be.SetCaret(at + 2, (be.LineAt(at + 2) ?? "").Length + 1);
                    be.InvalidateSnapshot();
                    var b = be.Snapshot();
                    bool jumped = a != null && b != null && b.Line != a.Line;
                    bool want = Rules.AutoTriggerWanted(true, false, true, jumped);
                    check("跳行不弹（鼠标单击 / 方向键）", jumped && !want,
                        "L" + (a?.Line ?? -1) + "->L" + (b?.Line ?? -1) +
                        " jumped=" + jumped + " want=" + want);
                    dropFrom();
                }
            }
            finally
            {
                // 删干净临时块（多删几次兜底），再逐字节比对原文
                for (int i = 0; i < 4; i++)
                {
                    int c = be.LineCount();
                    if (c <= baseCount) break;
                    be.DeleteLines(baseCount + 1, c - baseCount);
                }
            }

            string now2 = be.ModuleText();
            bool restored = string.Equals(now2, orig, StringComparison.Ordinal);
            if (!restored) allOk = false;
            sb.AppendLine("restored = " + restored + "   lines " + be.LineCount() + "/" + baseCount);
            sb.AppendLine(allOk ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "hktest.log"), sb.ToString(), Encoding.UTF8);
            return allOk ? 0 : 1;
        }
    }
}
