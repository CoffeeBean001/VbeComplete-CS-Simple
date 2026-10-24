// 换行/行操作命令 —— Ctrl+D / Ctrl+K / Ctrl+/ / Ctrl+Shift+Enter /
// Ctrl+↑↓（过程内的两个端点）。
//
// 为什么单独一个文件：这些是【会改用户代码】的动作，必须满足两条：
//   1. 同一语义只有一份实现 —— 钩子（App）和真机验证（Program --hktest）走同一批函数，
//      否则"验证通过"证明不了"按键时用的那条路是好的"之前就是这样，
//      验证和实际各写一份，结果光标错了两轮都没人发现）。
//   2. 返回值说清楚"到底做成了没有" —— 钩子靠它决定要不要把按键还回去。
//      旧写法是 void：拿不到快照就直接 return，键被吞掉、屏幕上一动不动，
//      用户只能得出"快捷键失效了"这个结论）。
//
// 全部只在 UI 线程调用（内部会走 COM）。

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal static class Commands
    {
        /// <summary>Ctrl+Shift+Enter：在当前行上方插一行，缩进跟着当前行走。</summary>
        public static bool InsertLineAbove(VbeBackend be)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;
            string indent = LeadingWs(snap.Text ?? "");
            if (!be.InsertLines(snap.Line, indent)) return false;
            be.SetCaret(snap.Line, indent.Length + 1);
            return true;
        }

        /// <summary>Ctrl+D：复制当前行到下一行，光标跟到副本上。</summary>
        public static bool DuplicateLine(VbeBackend be)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;
            string line = Rules.Rstrip(snap.Text ?? "");
            if (!be.InsertLines(snap.Line + 1, line)) return false;
            be.SetCaret(snap.Line + 1, CharColOf(snap, line));
            return true;
        }

        /// <summary>
        /// Ctrl+K：删掉当前整行，光标落到【上一行】的行尾。
        /// 删第 1 行 -> 新首行首列；模块只剩一行 -> 清空该行，照样算成功。
        /// </summary>
        public static bool DeleteLine(VbeBackend be)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;
            int count = be.LineCount();
            if (count <= 0) return false;

            if (count == 1) return be.ReplaceLineRaw(1, "");

            if (!be.DeleteLines(snap.Line, 1)) return false;

            if (snap.Line <= 1) be.SetCaret(1, 1);
            else be.CaretToLineEnd(snap.Line - 1);
            return true;
        }

        /// <summary>Ctrl+/：整行注释开关。有选区就逐行处理。</summary>
        public static bool ToggleComment(VbeBackend be)
        {
            long t0 = Environment.TickCount64;      // 从进函数算起：读选区+读行+执行都算
            // **不要用 Snapshot()**。它有一条"多行选区一律返回 null"的判据
            //   （见 VbeBackend.Snapshot 里 `if (sl != el) return null`），而 Ctrl+
            //   的典型用法恰恰是"选中一段代码批量注释"—— 于是用户每次按都被判成
            //   "拿不到快照"，键被吞掉、文档一个字不动，看起来就是"快捷键失效"。
            //   正确做法是直接 GetSelection + 逐行读写，绕开 Snapshot 的判据：
            //   选区自己取、行自己读、写完把选区原样摆回去。
            int sl, sc, el, ec;
            if (!be.TryGetSelection(out sl, out sc, out el, out ec)) return false;
            if (sl < 1 || el < sl) return false;

            // 选区跨行且停在**行首**（EndCol <= 1）：最后那一行其实没被真正选中，
            // 那是"连着换行符一起拖下来"的结果。要排除它。
            int from = sl, to = el;
            if (to > from && ec <= 1) to--;

            // 模块末尾那个"虚拟空行"（行号 = CountOfLines + 1）。夹一下，别越界去读。
            int count = be.LineCount();
            if (count > 0 && to > count) to = count;
            if (to < from) return false;

            // 整段**一次**读回来。原来是逐行 `be.LineAt(i)` —— 选 30 行就是
            //   30 次跨进程往返，光是"读"就够用户感觉到卡。VBE 的 `Lines(from, n)`
            //   一次就能把整段给你（CRLF 拼好），拆开用就行。
            int n = to - from + 1;
            string block = be.LinesBlock(from, n);
            if (block == null) return false;
            var lines = SplitBlock(block, n);
            if (lines.Count == 0) return false;

            // 全是注释 -> 取消；否则整段加注释。空行不参与"算不算注释"的判断。
            bool allCommented = true, anyContent = false;
            foreach (var l in lines)
            {
                string tr = l.TrimStart();
                if (tr.Length == 0) continue;
                anyContent = true;
                if (!tr.StartsWith("'", StringComparison.Ordinal)) { allCommented = false; break; }
            }
            if (!anyContent) return false;             // 全空白：根本没得注释，把键还回去

            // 先请 VBE 自己动手 —— `FindControl(1, 192|2552).Execute()` 等价于
            //   用户点"设置注释块 / 解除注释块"按钮。真机实测 3.5ms（手写逐行 3 行要
            //   15ms，30 行差得更远），而且：选区原样保留（不用再 SetSelection 摆回去）、
            //   改动进 VBE 自己的 Undo 栈（按一次 Ctrl+Z 整块回来）、空行处理与原生一致。
            //   只有它拿不到控件时（VBE 版本较旧 / 控件失效）才走下面的手写兜底。
            long ts = Environment.TickCount64;
            if (be.RunVbeCommand(allCommented
                    ? VbeBackend.CmdUncommentBlock
                    : VbeBackend.CmdCommentBlock))
            {
                // 这是 VBE 自己改的文档，我们这层的三份缓存（快照 / 模块全文 /
                //   作用域）一概不知道。原生命令会**保留选区**，行列都没变 ⇒ 下一次
                //   轮询的快照缓存照旧命中，Engine 拿"还少一个撇号"的旧行去算补全，
                //   按 Tab 就会吃掉/重复首字符。通知一次，全部作废重读。
                be.OnWeWrote();
                if (Log.Enabled)
                    Log.Write("Ctrl+/{0} {1} 行 @{2}：VBE 原生 Execute {3}ms，全程 {4}ms",
                        allCommented ? "解除" : "注释", lines.Count, from,
                        Environment.TickCount64 - ts, Environment.TickCount64 - t0);
                return true;
            }
            Log.Write("Ctrl+/{0} {1} 行 @{2}：VBE 原生命令不可用，回退逐行手写",
                allCommented ? "解除" : "注释", lines.Count, from);

            bool wrote = false;
            for (int i = 0; i < lines.Count; i++)
            {
                string l = lines[i];
                string nl;
                if (allCommented)
                {
                    int idx = l.IndexOf('\'');
                    nl = idx >= 0 ? l.Remove(idx, 1) : l;
                }
                else
                {
                    // 撇号加在**行首**（缩进之前），和 VBE 原生"注释块"按钮
                    //   一个字都不差。原来加在缩进之后（`    'n = 1`）——虽然语义一样，
                    //   但一旦原生路径可用/不可用来回切换，用户就会看到"有时 ' 在行首、
                    //   有时在缩进后"这种毫无道理的区别。两条路径的产出必须完全相同。
                    nl = "'" + l;
                }
                // 中途失败也继续写完：写一半停下会让文档卡在"部分注释"，比写完更糟。
                if (be.ReplaceLineRaw(from + i, nl)) wrote = true;
            }
            if (!wrote) return false;                   // 一行都没改成 -> 让调用方把键还回去

            // 摆光标 / 摆选区：行数没变，所以原来的行范围仍然有效。
            // 原来有选区就要**重新选中同样几行** —— 连按两次才能就地取消（IDEA 手感）。
            bool hadSel = sl != el || sc != ec;
            if (hadSel)
            {
                string last = lines[lines.Count - 1];
                int lastLen = allCommented ? Math.Max(1, last.Length) : last.Length + 1;
                // 列给足，让 VBE 自己钳到行尾 —— 绝不自己算显示列（铁律）
                be.SetSelectionRange(from, 1, to, lastLen + 100);
            }
            else
            {
                // 注释在行首插了一个 `'`、解除删掉了一个 —— 光标**前面**的
                //   字符数变了。原来直接把"改前"的显示列 `sc` 拿去换算，于是每按一次
                //   注释光标左移一格、每按一次解除右移一格（来回切换不守恒）。
                //   先按旧行算出旧字符列，再按增删补偏移，最后换回显示列。
                string oldLine = lines[0];
                int oldCharCol = Rules.CharColFromDisp(oldLine, Math.Max(1, sc));
                int newCharCol = Editing.AfterCommentCharCol(oldLine, oldCharCol, allCommented);
                string here = be.LineAt(from) ?? "";
                be.SetCaret(from, Rules.CaretColForCharCol(here, newCharCol));
            }
            return true;
        }

        /// <summary>
        /// `Ctrl+↑` / `Ctrl+↓`：过程内跳转。
        /// <para>↑：本过程里**最后一个 `Dim`** 语句的行尾（跨续行时落到语句结束行的行尾，
        /// 不会停在 `Dim a, _` 那一行）；↓：过程收尾行（`End Sub`）**上方最后一条代码**的行尾。</para>
        /// <para>光标不在任何过程里（模块声明区 / 两个过程之间 / 过程还没收尾）一律返回
        /// false —— 调用方把键原样还给系统（VBE 原生 `Ctrl+↑/↓` 是"跳到上/下一个过程"，
        /// 那是它自己的事）。判定不成立绝不吞键，这是全项目同一条规矩。</para>
        /// </summary>
        public static bool CtrlProcJump(VbeBackend be, int dir)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;
            string text = be.ModuleText();
            if (text == null) return false;

            var lines = Parser.SplitLines(text);
            var procs = Parser.ExtractProcedures(text);
            var proc = Parser.ProcOfLine(procs, snap.Line);
            if (proc == null) return false;              // 过程外：不接管

            int start = proc.StartLine, end = proc.EndLine;
            if (start < 1 || end > lines.Length || end <= start) return false;

            // 落点全在 Editing 的纯函数里（能被自检直接钉住），这里只管落光标。
            int target = dir < 0
                ? Editing.LastDimEndLine(lines, start, end)      // 最后一个 Dim 的语句结束行
                : Editing.ProcLastCodeLine(lines, start, end);   // 收尾行上方最后一条代码
            return target > 0 && be.CaretToLineEnd(target);
        }

        /// <summary>把"显示列"的光标位置折成同一行靠左的合理字符列。</summary>
        private static int CharColOf(Snapshot s, string line)
        {
            // 快照里的 CaretCol 是 VBE 的【显示列】（全角占 2 格、Tab 展开），
            //   不能当字符下标用 —— 含中文/Tab 的行上副本光标会偏到右边去。
            return Rules.CharColFromDisp(line, s.CaretCol);
        }

        internal static string LeadingWs(string s)
        {
            int i = 0;
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) i++;
            return s.Substring(0, i);
        }

        /// <summary>
        /// 把 `CodeModule.Lines(from, n)` 读回来的一整段拆成 n 行（行尾空白已去掉）。
        /// VBE 用 CRLF 拼接；末尾是空行时 split 会少一个元素，所以**按 n 补齐**——
        ///   "最后几行是空行"在注释块里极常见（选到 Sub 结尾那种），少一行就会
        ///   把上一行的撇号加到错的偏移上。
        /// </summary>
        private static List<string> SplitBlock(string block, int n)
        {
            var parts = (block ?? "").Split(new[] { "\r\n" }, StringSplitOptions.None);
            var list = new List<string>(n);
            for (int i = 0; i < n; i++)
                list.Add(i < parts.Length ? Rules.Rstrip(parts[i]) : "");
            return list;
        }
    }
}
