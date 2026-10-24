// 回车收尾的真机诊断：--enterprobe（会真的驱动一次键盘钩子）。
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
        /// 真机"回车补结构"探针 —— **走真键盘钩子**。
        ///
        /// 为什么非要有这一条：`--hktest` 是直接调 `VbeBackend.NewLineBelow`（等于
        ///   只验后端），而用户真正的路径是
        ///   `WH_KEYBOARD_LL → FreshHookSnap → EnterIndentWanted → Post → NewLineBelow`。
        ///   就是栽在这上面：后端 15 条全 PASS，真机上回车照样不补 `Next`
        ///   （用户第二次报同一个问题）。所以这里把真 `App` 起起来（真钩子 + 真轮询），
        ///   用 SendInput 敲键，再回读模块——**验的是用户手指走的那条路**。
        ///
        /// 用例：
        ///   A) 光标停在现成的 `    For i = 1 To 3` 行尾，按一下回车；
        ///   B) 从空行**逐字符敲** `    For i = 1 To 3`，最后一个字符后 0ms 立刻回车
        ///      （模拟真人手速：这条专门盯着 `_pendingKeys` 那个"输入法组字"判据）；
        ///   C) 同上敲 `    If x > 0 Then`，验 `End If`。
        /// </summary>
        private static int EnterProbe()
        {
            var sb = new StringBuilder();
            var gate = new object();
            bool allOk = true;
            Action<string, bool, string> check = (name, ok, detail) =>
            {
                if (!ok) allOk = false;
                lock (gate) sb.AppendLine((ok ? "  PASS  " : "  FAIL  ") + name + "   " + detail);
            };

            sb.AppendLine("=== enter probe（真机 / 走真钩子） ===");
            sb.AppendLine("（回车补结构 与 Ctrl+/ 注释都在这条真路径上验：钩子 -> 判据 -> Post -> COM）");

            var app = new App();
            var be = new VbeBackend();          // 主线程（STA）自己持有，不用 RunOnUi 绕
            int leftover = CleanupScratch(be);
            string orig = be.ModuleText();
            int baseCount = be.LineCount();

            sb.AppendLine("module lines = " + baseCount + "  len = " + (orig == null ? -1 : orig.Length) +
                          "  上次残留块清理 = " + leftover);
            // baseCount 可以是 0（空模块）：探针往里插脚手架再删干净，正好也能验
            if (baseCount < 0 || orig == null)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "enterprobe.log"), sb.ToString(), new UTF8Encoding(false));
                return 1;
            }

            // 造一段固定的脚手架：一个 Sub + 一行现成的 For + End Sub
            int at = baseCount + 1;
            RebuildScaffold(app, be, at, "    For i = 1 To 3");
            be.FocusVbe();

            // Start() 必须显式调用：装钩子 / 起托盘 / 起轮询定时器都在它里面
            app.Start();

            string diagA = "", diagA2 = "", diagB = "", diagC = "";
            int imeModeAtStart = -1;   // 探针把输入法切到英数档前记下原模式，跑完还回去

            var driver = new Thread(() =>
            {
                try
                {
                    // 1. 等轮询把"焦点在代码窗"跑出来 —— 钩子的第一道闸就是它。
                    //    从 bash / 计划任务起的进程 SetForegroundWindow 会被系统拒掉，
                    //      所以每轮都硬抢一次（AttachThreadInput 那套），抢到为止。
                    bool focused = false;
                    for (int i = 0; i < 40 && !focused; i++)
                    {
                        app.RunOnUi(() =>
                        {
                            Native.ForceForeground(Native.FindWindow("wndclass_desked_gsk", null));
                            be.SetCaret(at + 1, 19);
                            be.InvalidateSnapshot();
                        });
                        Thread.Sleep(250);
                        focused = app.CodePaneFocused;
                    }
                    check("焦点在代码窗（钩子的第一道闸）", focused, "CodePaneFocused=" + focused);
                    if (!focused) return;

                    // 输入法预检：B/C 是"逐字符打字"的用例，中文模式下每个字母都会被
                    // 输入法拿去组字（实测敲出来是 `Fori= 1 To3` / `If下》 0 Then`），
                    // 那不是程序坏了，是环境不对。这里直接把它切到英数档（走
                    // WM_IME_CONTROL），跑完再恢复成用户原来的模式。
                    int imeMode0 = Native.ImeConversionMode(Native.FocusedVbeWindow());
                    bool alnum0 = Native.ImeAlphanumeric(Native.FocusedVbeWindow());
                    if (!alnum0)
                    {
                        Native.ImeSetAlphanumeric(Native.FocusedVbeWindow());
                        Thread.Sleep(200);
                    }
                    bool alnum1 = Native.ImeAlphanumeric(Native.FocusedVbeWindow());
                    imeModeAtStart = imeMode0;
                    lock (gate)
                        sb.AppendLine("IME 英数模式 = " + alnum0 + " -> " + alnum1 +
                            "（conv=0x" + imeMode0.ToString("X2") + "，跑完恢复）" +
                            (alnum1 ? "" : "   ⚠️ 切不过去：B/C 会被输入法吃掉"));

                    // 2. 用例 A：现成的行，下面紧跟 End Sub，直接回车
                    Thread.Sleep(300);
                    diagA = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] a = ReadLines(app, be, at, 4);
                    check("A 现成 For 行 + 回车 -> 补 Next", a[3] == "    Next",
                        Tail(a) + "  diag{" + diagA + "}");

                    // 3. 用例 A2：过程中间插一个 For，
                    //    下面压着同级的语句 —— 必须补 Next。
                    RebuildScaffold(app, be, at, "    For i = 1 To 3", "    x = 1", "    y = 2");
                    Thread.Sleep(300);
                    diagA2 = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] a2 = ReadLines(app, be, at, 6);
                    check("A2 同级已有代码压着也要补 Next",
                        a2[3] == "    Next" && a2[4] == "    x = 1",
                        "[" + a2[1] + "][" + a2[2] + "][" + a2[3] + "][" + a2[4] + "]  diag{" + diagA2 + "}");

                    // 4. 用例 B：空行上逐字符敲 For 行（模拟真人），最后一下立刻回车
                    RebuildScaffold(app, be, at, "");
                    Thread.Sleep(300);
                    bool typedB = TypeTyping(app, "    For i = 1 To 3", 45, out string howB);
                    diagB = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] b = ReadLines(app, be, at, 4);
                    check("B 逐字符敲 For + 立刻回车 -> 补 Next", typedB && b[3] == "    Next",
                        "typed=" + typedB + " via=" + howB + " " + Tail(b) + "  diag{" + diagB + "}");

                    // 5. 用例 C：If … Then -> End If
                    RebuildScaffold(app, be, at, "");
                    Thread.Sleep(300);
                    bool typedC = TypeTyping(app, "    If x > 0 Then", 45, out string howC);
                    diagC = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] c = ReadLines(app, be, at, 4);
                    check("C 逐字符敲 If … Then -> 补 End If", typedC && c[3] == "    End If",
                        "typed=" + typedC + " via=" + howC + " " + Tail(c) + "  diag{" + diagC + "}");

                    // 6. 用例 D：单选 If（`If x Then y = 1`）—— 不该补 End If，
                    //    但**回车必须照常换行**（这一档走的是"先接管、判不过再还回去"，
                    //    最容易写出"吞了却什么都没发生"的就是这条路）。
                    RebuildScaffold(app, be, at, "    If x Then y = 1");
                    int beforeD = 0;
                    app.RunOnUi(() => { be.ForceReloadModule(); beforeD = be.LineCount(); });
                    Thread.Sleep(300);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(800);
                    string[] d = ReadLines(app, be, at, 4);
                    int afterD = 0;
                    app.RunOnUi(() => { be.ForceReloadModule(); afterD = be.LineCount(); });
                    check("D 单选 If 回车：不补 End If 但必须换行",
                        afterD == beforeD + 1 && d[1] == "    If x Then y = 1" && d[2].Trim() == "",
                        "lines " + beforeD + "->" + afterD + " " + Tail(d));

                    // 7. 用例 E：Ctrl+/ 注释开关 —— 走【同一条真键盘路径】。
                    //    命令本身另有 --hktest 直接调 Commands 验过；这里要验的是
                    //      "钩子有没有把 Ctrl+/ 路由进那个分支、有没有把键吞了却没做事"。
                    //      "Ctrl+/ 失效"时，只有这条路能证明卡在哪。
                    RebuildScaffold(app, be, at, "    n = 1");
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(700);
                    string[] e1 = ReadLines(app, be, at, 3);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(700);
                    string[] e2 = ReadLines(app, be, at, 3);
                    check("E Ctrl+/ 整行注释开关（真钩子路径）",
                        e1[1].TrimStart().StartsWith("'", StringComparison.Ordinal)
                        && e1[1].Contains("n = 1")
                        && e2[1] == "    n = 1",
                        "加后=[" + e1[1] + "] 还原后=[" + e2[1] + "] want=[    n = 1]");

                    // 8. 用例 F：**多行选区**批量注释。
                    //    注意不能走 `be.Snapshot()`：Snapshot 对多行选区一律返回
                    //    null（VbeBackend 里 `if (sl != el) return null`）⇒ 每次按都判成
                    //    "拿不到快照"，键被吞、文档一个字不动。单选却一直是好的，所以
                    //    这条必须单独验：选中两行 -> Ctrl+/ -> 两行都该加上撇号；
                    //    再按一次（选区会被我们原样摆回去）-> 两行都该还原。
                    RebuildScaffold(app, be, at, "    n = 1", "    n = 2");
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.SetSelectionRange(at + 1, 1, at + 2, 1 + "    n = 2".Length + 100);
                    });
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(800);
                    string[] f1 = ReadLines(app, be, at, 4);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(800);
                    string[] f2 = ReadLines(app, be, at, 4);
                    check("F 多行选区 Ctrl+/ 批量注释",
                        f1[1].TrimStart().StartsWith("'", StringComparison.Ordinal)
                        && f1[2].TrimStart().StartsWith("'", StringComparison.Ordinal)
                        && f2[1] == "    n = 1" && f2[2] == "    n = 2",
                        "加后=[" + f1[1] + "][" + f1[2] + "] 还原后=[" + f2[1] + "][" + f2[2] + "]");

                    // 9. 用例 G：**40 行**批量注释 —— 验 换上的"VBE 原生执行"
                    //    在大选区上既不慢也不错。
                    //    逐行写法：40 次 Lines(i,1) + 40 次 ReplaceLine + 1 次 SetSelection
                    //    = 81 次跨进程往返（真机 340ms 上下，用户就是这么感觉到"慢"的）；
                    //    现在：1 次 Lines(from,n) + 1 次 Execute + 0 次还原（选区原生保留）。
                    var many = new string[39];
                    for (int i = 1; i <= 39; i++) many[i - 1] = "    v" + i + " = " + i;
                    RebuildScaffold(app, be, at, "    v0 = 0", many);
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.SetSelectionRange(at + 1, 1, at + 40, 200);
                    });
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(900);
                    string[] g1 = ReadLines(app, be, at, 42);
                    int quoted = 0;
                    for (int i = 1; i <= 40; i++)
                        if (g1[i].TrimStart().StartsWith("'", StringComparison.Ordinal)) quoted++;
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(900);
                    string[] g2 = ReadLines(app, be, at, 42);
                    int plain = 0;
                    for (int i = 1; i <= 40; i++)
                        if (!g2[i].TrimStart().StartsWith("'", StringComparison.Ordinal) && g2[i].Contains("v"))
                            plain++;

                    // 同一批 40 行，直接量"命令本身"的耗时（不掺键盘延迟）。
                    long msSet = -1, msUnset = -1;
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.SetSelectionRange(at + 1, 1, at + 40, 200);
                        long t = Environment.TickCount64;
                        Commands.ToggleComment(be);
                        msSet = Environment.TickCount64 - t;
                        long t2 = Environment.TickCount64;
                        Commands.ToggleComment(be);
                        msUnset = Environment.TickCount64 - t2;
                    });

                    check("G 40 行批量注释（一次 Execute，不再逐行）",
                        quoted == 40 && plain == 40 && msSet >= 0 && msSet < 200
                        && g1[40].TrimStart().StartsWith("'", StringComparison.Ordinal),
                        "注释 " + quoted + "/40 · 还原 " + plain + "/40 · 耗时 "
                        + msSet + "ms/" + msUnset + "ms（逐行写法约 340ms）");

                    // 10. 用例 H：**空格收起候选窗** —— "set + 空格不收、let 收"。
                    //     钩子若对空格什么都不做、全指望轮询的 changed 兜，则
                    //     changed 比的是 **Rstrip 之后**的文本（行尾空白刻意不算文档
                    //     变化），`Set ` 与 `Set` 在它眼里一字不差 ⇒ 永远走不到 Trigger
                    //     ⇒ 列表就那么挂着（轮询那条 else 只是把窗子挪到新光标位置）。
                    //     let 之所以"看起来对"，是它打到第三个字母时已判成"打全了"
                    //     提前收的，跟空格无关。
                    //
                    //     这里用 Ctrl+Space 手动唤出（比"打字才弹"稳定，不受输入法
                    //     中/英状态影响），再发一个**裸空格**：既要验列表收了，也要验
                    //     空格照常落进文档（我们只收窗、不吞键）。
                    RebuildScaffold(app, be, at, "    n = 1");
                    app.RunOnUi(() => { be.ForceReloadModule(); be.SetCaret(at + 1, 1); });
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_SPACE, ctrl: true);
                    Thread.Sleep(600);
                    bool shownH = app.PopupVisible;
                    SendKeys(app, Native.VK_SPACE);
                    Thread.Sleep(500);
                    bool hiddenH = !app.PopupVisible;
                    string[] h = ReadLines(app, be, at, 3);
                    check("H 空格收起候选窗（set + 空格不收）",
                        shownH && hiddenH && h[1] == "     n = 1",
                        "唤出=" + shownH + " 空格后收起=" + hiddenH + " 行=[" + h[1] + "]");

                    // 11. 用例 I：**Ctrl+↑ / Ctrl+↓（不带 Shift）** —— 两个过程内跳转。
                    //     它们与**带 Shift** 的那一组语义完全不同（↑ = 光标上方最近的声明行、
                    //     ↓ = 下方第一个空白行）；不带 Shift 的这两下是本过程的端点，
                    //     不能漏给 VBE 原生（那会跳到上/下一个过程）。
                    //     这条同时能证明"键到底有没有被路由进新分支"：判据不成立时会把键
                    //     还回去（Ctrl 被物理按住 => 实际发出裸 ↑/↓），光标只会走一行，
                    //     断言必然对不上。
                    RebuildScaffold(app, be, at, "    x = 1",
                        "    Dim a As Long", "    y = 2", "    Dim b As Long", "    z = 3");
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_UP, ctrl: true);
                    Thread.Sleep(600);
                    int upLine = 0, upCol = 0;
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.TryGetSelection(out int l1, out int c1, out _, out _);
                        upLine = l1; upCol = c1;
                    });
                    SendCombo(app, Native.VK_DOWN, ctrl: true);
                    Thread.Sleep(600);
                    int dnLine = 0, dnCol = 0;
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.TryGetSelection(out int l2, out int c2, out _, out _);
                        dnLine = l2; dnCol = c2;
                    });
                    const string lastDimLine = "    Dim b As Long";   // at+4
                    const string lastCodeLine = "    z = 3";          // at+5
                    check("I Ctrl+↑/↓ 跳过程内两个端点",
                        upLine == at + 4 && upCol == lastDimLine.Length + 1
                        && dnLine == at + 5 && dnCol == lastCodeLine.Length + 1,
                        "↑=" + upLine + "," + upCol + " want " + (at + 4) + "," + (lastDimLine.Length + 1)
                        + "  ↓=" + dnLine + "," + dnCol + " want " + (at + 5) + "," + (lastCodeLine.Length + 1));
                }
                catch (Exception e)
                {
                    lock (gate) sb.AppendLine("EXCEPTION " + e);
                    allOk = false;
                }
                finally
                {
                    // 把输入法还给用户原来的状态（探针把它切到英数档才能逐字符打字）
                    Native.ImeRestore(Native.FocusedVbeWindow(), imeModeAtStart);
                    app.RequestExit();
                }
            })
            { IsBackground = true, Name = "enterprobe" };

            driver.Start();
            // 整段清理放进 finally：探针中途抛异常（COM 抖动、脚手架建失败）时
            //   也绝不能把 `Sub VbcEnter()` 留在用户模块里。只在正常路径清理的话，
            //   一旦半路炸掉，垃圾就永久留在用户的代码里。
            bool joined = false;
            try
            {
                Application.Run(app);
            }
            finally
            {
                try
                {
                    joined = driver.Join(4000);
                    // 这里不能再用 app.RunOnUi：Application.Run 已经返回，_disp 的
                    //   窗口没了，Invoke 会抛（被吞掉）—— 上一次就是因此把脚手架
                    //   留在了用户模块里。
                    be.ForceReloadModule();
                    int removed = CleanupScratch(be);
                    be.ForceReloadModule();
                    string now = be.ModuleText();
                    bool restored = string.Equals(now, orig, StringComparison.Ordinal);
                    if (!restored) allOk = false;
                    lock (gate)
                    {
                        sb.AppendLine("cleanup 块 = " + removed + "   restored = " + restored +
                                      "   lines " + be.LineCount() + "/" + baseCount +
                                      (joined ? "" : "   ⚠️ driver 线程超时未结束"));
                        sb.AppendLine(allOk ? "RESULT: PASS" : "RESULT: FAIL");
                        File.WriteAllText(Path.Combine(Settings.Home, "enterprobe.log"),
                            sb.ToString(), new UTF8Encoding(false));
                    }
                }
                catch (Exception e)
                {
                    // 清理本身失败：至少把"没清干净"喊出来，别让人以为模块是干净的
                    lock (gate)
                    {
                        sb.AppendLine("cleanup 异常：" + e.Message);
                        sb.AppendLine("RESULT: FAIL");
                        try
                        {
                            File.WriteAllText(Path.Combine(Settings.Home, "enterprobe.log"),
                                sb.ToString(), new UTF8Encoding(false));
                        }
                        catch { }
                    }
                    allOk = false;
                }
            }
            return allOk ? 0 : 1;
        }
    }
}
