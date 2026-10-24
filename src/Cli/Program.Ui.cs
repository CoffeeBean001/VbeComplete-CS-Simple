// 离屏界面预览与探针：--uitest、--probe、--yieldprobe。
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
        /// 候选窗离屏渲染：把窗口画进一张 PNG，用来一眼核对配色 / 图标 / 布局。
        /// 不需要桌面，也不会出现在屏幕上。
        /// </summary>
        private static int UiTest()
        {
            var sb = new StringBuilder();
            try
            {
                var f = new PopupForm();
                // 用**真实匹配 + 真实排序**造数据：预览图里看到的高亮位置、
                // 图标档位、行序，都与用户在 VBE 里看到的一致。
                // 手写 Pos/Span 的那种预览图，改了算法它也不会跟着变。
                // 池子里只剩自己写的名字，档位（func / prop / var）
                //   用注入的一段合成源码算出来。
                var eng = new Engine { Recent = new Recent() };
                eng.PrimePool(
                    "Public Sub SelectAll()\r\nEnd Sub\r\n" +
                    "Public Function selBegin() As Long\r\nEnd Function\r\n" +
                    "Public Property Get Size() As Long\r\nEnd Property\r\n" +
                    "Private mCount As Long\r\n",
                    "Module1");
                const string word = "se";
                var cands = new System.Collections.Generic.List<Candidate>();
                foreach (var n in new[]
                         {
                             "SelectAll", "selBegin", "Size", "dataSheet",
                             "studentName", "mCount", "设置表格",
                         })
                {
                    FuzzyHit fh;
                    if (!Fuzzy.Match(n, word, out fh)) continue;
                    cands.Add(new Candidate
                    {
                        Name = n,
                        Kind = eng.Classify(n),
                        Pos = fh.First,
                        Span = fh.Span,
                        Hits = fh.Positions,
                        HitKind = fh.Kind,
                        ScopeKind = 3,
                    });
                }
                cands.Sort(Engine.Cmp);

                f.SetData(cands, 1, 0, 15);
                var size = f.DesiredSize();
                int w = size.Width;
                int h = size.Height;
                f.StartPosition = FormStartPosition.Manual;
                f.Location = new System.Drawing.Point(-4000, -4000);
                f.Show();
                f.Width = w; f.Height = h;
                Application.DoEvents();

                using (var bmp = new System.Drawing.Bitmap(w, h))
                {
                    f.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, w, h));
                    string path = Path.Combine(Settings.Home, "popup-preview.png");
                    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                    sb.AppendLine("saved " + path + "  " + w + "x" + h +
                                  "  items=" + cands.Count + "  word=" + word +
                                  "  font=" + f.FontName + "  rowH=" + f.RowHeight);
                }
                f.Hide();
                f.Dispose();

                // ---- 片段窗也离屏渲染一张：布局/配色/按钮顺序一眼能核对 ----
                try
                {
                    using (var store = new SnippetStore())
                    {
                        var sf = new SnippetForm(store, code => false);
                        sf.StartPosition = FormStartPosition.Manual;
                        sf.Location = new System.Drawing.Point(-4000, -4000);
                        sf.Show();
                        sf.PreviewOnly(store.All());
                        Application.DoEvents();
                        int sw = sf.Width, sh = sf.Height;
                        using (var bmp2 = new System.Drawing.Bitmap(sw, sh))
                        {
                            sf.DrawToBitmap(bmp2, new System.Drawing.Rectangle(0, 0, sw, sh));
                            string p2 = Path.Combine(Settings.Home, "snippet-preview.png");
                            bmp2.Save(p2, System.Drawing.Imaging.ImageFormat.Png);
                            sb.AppendLine("saved " + p2 + "  " + sw + "x" + sh +
                                           "  rows=" + store.Count());
                        }
                        sf.Hide();
                        sf.Dispose();
                    }
                }
                catch (Exception e2)
                {
                    sb.AppendLine("snippet preview skipped: " + e2.Message);
                }

                sb.AppendLine("RESULT: PASS");
                File.WriteAllText(Path.Combine(Settings.Home, "uitest.log"), sb.ToString(), Encoding.UTF8);
                return 0;
            }
            catch (Exception e)
            {
                sb.AppendLine("EXCEPTION " + e);
                File.WriteAllText(Path.Combine(Settings.Home, "uitest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }
        }

        /// <summary>诊断用：把"能不能连上 VBE"写进 probe.log。</summary>
        private static int Probe()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Home = " + Settings.Home);
            sb.AppendLine("settings.ini = " + Path.Combine(Settings.Home, "settings.ini"));
            sb.AppendLine("settings error = " + Settings.LastError());
            sb.AppendLine("kernel = " + Environment.Version);
            sb.AppendLine("64bit = " + Environment.Is64BitProcess);

            var be = new VbeBackend();
            sb.AppendLine("--- COM ---");
            foreach (var progid in new[] { "Excel.Application", "Word.Application" })
            {
                object app = ComCtx.GetActive(progid);
                sb.AppendLine(progid + " active = " + (app == null ? "null" : "ok"));
                if (app == null) continue;
                try
                {
                    dynamic dapp = app;
                    dynamic vbe = dapp.VBE;
                    sb.AppendLine("  VBE = " + (vbe == null ? "null" : "ok"));
                    if (vbe == null) continue;
                    try
                    {
                        dynamic projs = vbe.VBProjects;
                        int pc = Convert.ToInt32(projs.Count);
                        sb.AppendLine("  projects = " + pc);
                        if (pc > 0)
                        {
                            dynamic p0 = projs.Item(1);
                            sb.AppendLine("  project[1].Name = " + Convert.ToString(p0.Name));
                            dynamic comps = p0.VBComponents;
                            sb.AppendLine("  components = " + Convert.ToInt32(comps.Count));
                        }
                    }
                    catch (Exception e) { sb.AppendLine("  projects ERR " + e.Message); }
                    try
                    {
                        dynamic cp = vbe.ActiveCodePane;
                        sb.AppendLine("  ActiveCodePane = " + (cp == null ? "null" : "ok"));
                        if (cp != null)
                        {
                            dynamic cm = cp.CodeModule;
                            sb.AppendLine("  module = " + Convert.ToString(cm.Name) +
                                          " lines = " + Convert.ToInt32(cm.CountOfLines));
                            try
                            {
                                dynamic sel = cp.Selection;
                                sb.AppendLine("  Selection = " +
                                    Convert.ToInt32(sel.StartLine) + "," + Convert.ToInt32(sel.StartColumn) +
                                    " .. " + Convert.ToInt32(sel.EndLine) + "," + Convert.ToInt32(sel.EndColumn));
                            }
                            catch (Exception e) { sb.AppendLine("  Selection ERR " + e.Message); }
                            try
                            {
                                object g = cp.GetSelection();
                                sb.AppendLine("  GetSelection = " + (g == null ? "null" : g.GetType().FullName));
                            }
                            catch (Exception e) { sb.AppendLine("  GetSelection ERR " + e.Message); }
                        }
                    }
                    catch (Exception e) { sb.AppendLine("  ActiveCodePane ERR " + e.Message); }
                }
                catch (Exception e) { sb.AppendLine("  VBE ERR " + e.Message); }
            }

            sb.AppendLine("--- backend ---");
            var snap = be.Snapshot();
            sb.AppendLine("snapshot = " + (snap == null ? "null" : snap.ToString()));
            if (snap != null)
            {
                sb.AppendLine("line text = " + (snap.Text ?? ""));
                sb.AppendLine("module text len = " + (be.ModuleText()?.Length ?? -1));
                sb.AppendLine("components = " + string.Join(", ", be.ComponentNames()));
            }
            sb.AppendLine("--- settings ---");
            sb.AppendLine("NO_AUTOSPACE = " + Settings.Flag("NO_AUTOSPACE", false));
            sb.AppendLine("NO_AUTOPAREN = " + Settings.Flag("NO_AUTOPAREN", false));
            sb.AppendLine("HIDE_ON_COMPLETE = " + Settings.Flag("HIDE_ON_COMPLETE", true));
            sb.AppendLine("VIEW_ROWS = " + Settings.Int("VIEW_ROWS", 15, 3, 30));
            sb.AppendLine("--- win32 ---");
            sb.AppendLine("inVbeCodePane(win32) = " + DescribeVbeFocus());
            // 让位判据（App.PollOnce 每轮都问这一个函数）
            try
            {
                sb.AppendLine("vbeMemberListVisible = " + Native.VbeMemberListVisible());
                var rects = Native.VbePopupRects();
                sb.AppendLine("vbePopupRects = " + (rects == null || rects.Length == 0 ? "(none)" :
                    string.Join(" | ", rects.Select(r => r.Left + "," + r.Top + "-" + r.Right + "," + r.Bottom))));
            }
            catch (Exception e) { sb.AppendLine("vbeMemberListVisible ERR " + e.Message); }
            File.WriteAllText(Path.Combine(Settings.Home, "probe-cs.log"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

        /// <summary>
        /// 只采样一次"让位判据"（App.PollOnce 用的同一个函数）。
        /// 不碰 COM，一次约 2ms —— 真机上要连测"VBE 的列表出现时我们能不能看见"，
        /// 走 --probe 太重（要过一遍 COM，几秒钟）。
        /// </summary>
        private static int YieldProbe()
        {
            var sb = new StringBuilder();
            int visible = 0, seen = 0;
            for (int i = 0; i < 5; i++)
            {
                seen++;
                bool vis = Native.VbeMemberListVisible();
                if (vis) visible++;
                var rects = Native.VbePopupRects();
                sb.AppendLine(DateTime.Now.ToString("HH:mm:ss.fff") +
                              " memberListVisible=" + vis +
                              " rects=" + (rects == null || rects.Length == 0 ? "(none)" :
                                  string.Join(" | ", rects.Select(r => r.Left + "," + r.Top + "-" + r.Right + "," + r.Bottom))));
                Thread.Sleep(200);      // 缓存 150ms，睡久一点才能读到新采样
            }
            sb.AppendLine("RESULT: " + (visible > 0 ? "VBE-LIST-VISIBLE" : "VBE-LIST-NOT-VISIBLE") +
                          " (" + visible + "/" + seen + ")");
            File.WriteAllText(Path.Combine(Settings.Home, "yieldprobe.log"), sb.ToString(), Encoding.UTF8);
            return visible > 0 ? 0 : 1;
        }

        /// <summary>VBE 主框架窗口 + 当前焦点窗口的类名（探针里用来核对"焦点在不在代码窗"）。</summary>
        private static string DescribeVbeFocus()
        {
            IntPtr frame = Native.FindWindow("wndclass_desked_gsk", null);
            if (frame == IntPtr.Zero) return "no vbe frame";
            uint pid;
            uint tid = Native.GetWindowThreadProcessId(frame, out pid);
            IntPtr f = Native.FocusedHwndOf(tid);
            return "frame=" + frame + " pid=" + pid + " focusClass=" + Native.ClassNameOf(f);
        }
    }
}
