// 候选池 / 模块全文的离线诊断：--dump、--pooldump。
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
        /// 把当前代码窗模块的全文写到 moddump.txt（只读，一个字节都不改）。
        /// 真机跑 --writetest / --hktest / --carettest 之前先 dump 一份，
        /// 跑完再 dump 一份 diff —— 那几条用例都会"改完再删"，这是唯一的核对手段。
        /// </summary>
        private static int DumpModule()
        {
            var be = new VbeBackend();
            string text = be.ModuleText();
            var snap = be.Snapshot();
            var sb = new StringBuilder();
            sb.AppendLine("module = " + (snap == null ? "?" : snap.Module) +
                          "  lines = " + be.LineCount() +
                          "  caret = " + (snap == null ? "?" : snap.Line + ":" + snap.CaretCol));
            sb.AppendLine("----");
            sb.Append(text ?? "(null)");
            File.WriteAllText(Path.Combine(Settings.Home, "moddump.txt"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

        /// <summary>
        /// 诊断：把当前模块的解析结果 + 光标处的候选全写进 pooldump.txt。
        /// <para>"某个名字该提示却没提示"时，跑一次这个，
        /// 一次就能分清是"没进 FreeIdents"、"进了但作用域对不上"、还是
        /// "进了但被 Option Explicit 拦了"。比来回猜快得多。</para>
        /// 全程只读，不改任何代码。
        /// </summary>
        private static int PoolDump(string onlyModule)
        {
            var sb = new StringBuilder();
            // 带模块名时：读那个模块的全文做"离线解析"（不跑 Trigger，只回答
            // "这个名字到底进没进 FreeIdents / 挂在哪个过程"）。
            if (!string.IsNullOrEmpty(onlyModule)) return PoolDumpNamed(onlyModule, sb);
            try
            {
                var be = new VbeBackend();
                string text = be.ModuleText();
                var snap = be.Snapshot();
                if (snap == null || text == null)
                {
                    sb.AppendLine("no code module（VBE 里没有打开任何代码窗）");
                    File.WriteAllText(Path.Combine(Settings.Home, "pooldump.txt"), sb.ToString(), Encoding.UTF8);
                    return 1;
                }

                var ms = Scope.Analyze(snap.Module, text);
                var proc = Parser.ProcOfLine(ms.Procs, snap.Line);
                string procName = proc == null ? "(null)" : proc.ScopeKey;

                sb.AppendLine("=== pool dump ===");
                sb.AppendLine("module   = " + snap.Module + "  lines = " + be.LineCount());
                sb.AppendLine("optionExplicit = " + ms.OptionExplicit);
                sb.AppendLine("caret    = L" + snap.Line + " C" + snap.CaretCol + "  line = [" + snap.Text + "]");
                sb.AppendLine("curProc  = " + procName);
                sb.AppendLine("procs    = " + string.Join(", ",
                    ms.Procs.Select(p => p.Name + ":" + p.StartLine + "-" + p.EndLine + "(" + p.ScopeKey + ")")));
                sb.AppendLine("decls    = " + string.Join(", ",
                    ms.Decls.Select(d => d.Name + "@" + (d.Proc ?? "-"))));
                sb.AppendLine("free     = " + string.Join(", ", ms.FreeIdents.Select(f =>
                {
                    HashSet<string> ps;
                    ms.FreeIdentProc.TryGetValue(f, out ps);
                    string where = ps == null || ps.Count == 0
                        ? "-"
                        : string.Join("/", ps.Select(x => x.Length == 0 ? "(模块级)" : x));
                    return f + "@" + where;
                })));
                sb.AppendLine("visible  = " + string.Join(", ", ms.Visible(procName).Keys.Take(80)));

                var eng = new Engine { Backend = be };
                bool shown = eng.Trigger(false, snap);
                sb.AppendLine("trigger  = " + shown + "  items = " + eng.Items.Count +
                              (eng.TruncatedFrom > 0 ? " (total " + eng.TruncatedFrom + ")" : ""));
                sb.AppendLine("cands    = " + string.Join(", ", eng.Items.Take(40).Select(c => c.Name)));
            }
            catch (Exception e)
            {
                sb.AppendLine("EXCEPTION " + e);
            }
            File.WriteAllText(Path.Combine(Settings.Home, "pooldump.txt"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

        /// <summary>按模块名离线解析（不需要那个代码窗是活动的）。</summary>
        private static int PoolDumpNamed(string name, StringBuilder sb)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) { sb.AppendLine("no VBE"); return 1; }
                dynamic comps = vbe.ActiveVBProject.VBComponents;
                int n = Convert.ToInt32(comps.Count);
                for (int i = 1; i <= n; i++)
                {
                    dynamic c = comps.Item(i);
                    string nm = Convert.ToString(c.Name) ?? "";
                    if (!string.Equals(nm, name, StringComparison.OrdinalIgnoreCase)) continue;
                    dynamic cm = c.CodeModule;
                    int cnt = Convert.ToInt32(cm.CountOfLines);
                    string text = cnt > 0 ? Convert.ToString(cm.Lines(1, cnt)) : "";
                    var ms = Scope.Analyze(nm, text ?? "");
                    sb.AppendLine("=== pool dump (module " + nm + ", lines " + cnt + ") ===");
                    sb.AppendLine("optionExplicit = " + ms.OptionExplicit);
                    sb.AppendLine("procs    = " + string.Join(", ",
                        ms.Procs.Select(p => p.Name + ":" + p.StartLine + "-" + p.EndLine + "(" + p.ScopeKey + ")")));
                    sb.AppendLine("decls    = " + string.Join(", ",
                        ms.Decls.Select(d => d.Name + "@" + (d.Proc ?? "-"))));
                    sb.AppendLine("free     = " + string.Join(", ", ms.FreeIdents.Select(f =>
                    {
                        HashSet<string> ps;
                        ms.FreeIdentProc.TryGetValue(f, out ps);
                        string where = ps == null || ps.Count == 0
                            ? "-"
                            : string.Join("/", ps.Select(x => x.Length == 0 ? "(模块级)" : x));
                        return f + "@" + where;
                    })));
                    File.WriteAllText(Path.Combine(Settings.Home, "pooldump.txt"), sb.ToString(), Encoding.UTF8);
                    return 0;
                }
                sb.AppendLine("module not found: " + name);
            }
            catch (Exception e) { sb.AppendLine("EXCEPTION " + e); }
            File.WriteAllText(Path.Combine(Settings.Home, "pooldump.txt"), sb.ToString(), Encoding.UTF8);
            return 1;
        }
    }
}
