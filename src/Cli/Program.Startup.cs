// 启动环境诊断：--startuptest（内嵌图标 + 开机自启项）。
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
        /// 图标 + 开机自启的自检（不需要 Excel、不需要桌面）。
        ///   · 图标：内嵌资源在不在、九个尺寸能不能各自取到、拿到的像素尺寸对不对；
        ///   · 自启：登记 → 读回来要等于当前命令行 → 取消 → 读回来要没有。
        /// 先给注册表和 autostart.txt 拍快照，测完原样还原 —— 不留任何痕迹。
        /// </summary>
        private static int StartupTest()
        {
            var sb = new StringBuilder();
            bool ok = true;
            Action<string, bool, string> check = (name, pass, detail) =>
            {
                if (!pass) ok = false;
                sb.AppendLine((pass ? "  PASS  " : "  FAIL  ") + name +
                              (string.IsNullOrEmpty(detail) ? "" : "   " + detail));
            };

            sb.AppendLine("=== icon + autostart test ===");

            // ---- 1. 图标 ----
            var asm = typeof(IconAssets).Assembly;
            bool hasRes = false;
            foreach (var n in asm.GetManifestResourceNames())
                if (n.EndsWith("app.ico", StringComparison.OrdinalIgnoreCase)) hasRes = true;
            check("内嵌 app.ico 资源存在", hasRes,
                "resources=" + string.Join(",", asm.GetManifestResourceNames()));

            int[] sizes = { 16, 20, 24, 32, 40, 48, 64, 128, 256 };
            var bad = new StringBuilder();
            foreach (int s in sizes)
            {
                try
                {
                    // 共享实例，不能 using —— 否则跑完检查就把缓存里的图标 Dispose 了
                    {
                        var ic = IconAssets.Get(s);
                        // Icon.Width 报的是"选中那一帧"的尺寸，可能不等于请求值。
                        // 唯一已知的例外：256 那一帧是 PNG 压缩的（ICO 规范允许），
                        //   .NET 的 Icon(stream, size) 挑不到它，会退回 128 —— 我们
                        //   实际只用 16/20/32/48，256 是给资源管理器看 exe 用的，
                        //   所以这里只要求"别比请求的小太多"。
                        int want = Math.Min(s, 128);
                        if (ic == null || ic.Width < want || ic.Width > s)
                            bad.Append(s + "->" + (ic == null ? "null" : ic.Width.ToString()) + " ");
                    }
                }
                catch (Exception e) { bad.Append(s + "->异常:" + e.Message + " "); }
            }
            check("九个尺寸都能取到图标", bad.Length == 0, bad.ToString());

            // 托盘那一份必须不是"通用空白图标"（和 SystemIcons.Application 尺寸相同但内容不同，
            // 这里只做存在性判断）
            var ti = IconAssets.Tray();     // 共享实例，不 Dispose
            check("托盘图标可用", ti != null && ti.Width >= 16, "w=" + (ti == null ? -1 : ti.Width));
            check("exe 路径可取", !string.IsNullOrEmpty(AutoStart.ExePath()), AutoStart.ExePath());

            // 把真正会挂到托盘上的那两个尺寸存成图片，肉眼确认"是不是那张蓝色 V"
            // （只看 log 里"加载成功"是不够的 —— 加载成功也可能是退回系统图标）
            try
            {
                // Icon 是共享实例（不 Dispose）；只有 ToBitmap 出来的 Bitmap 归我们。
                var i32 = IconAssets.Get(32);
                using (var b32 = i32.ToBitmap())
                using (var canvas = new Bitmap(64 * 4 + 24, 80))
                {
                    using (var g = Graphics.FromImage(canvas))
                    {
                        g.Clear(Color.FromArgb(32, 34, 38));
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                        g.DrawImage(b32, new Rectangle(8, 8, 32, 32));
                        g.DrawImage(b32, new Rectangle(48, 8, 32, 32));
                        var i16 = IconAssets.Get(16);
                        using (var b16 = i16.ToBitmap())
                            g.DrawImage(b16, new Rectangle(88, 16, 16, 16));
                    }
                    canvas.Save(Path.Combine(Settings.Home, "icon-tray.png"),
                        System.Drawing.Imaging.ImageFormat.Png);
                }
            }
            catch (Exception e) { sb.AppendLine("  存 icon-tray.png 失败：" + e.Message); }

            // ---- 2. 开机自启 ----
            var snap = AutoStart.Capture();
            try
            {
                sb.AppendLine("  原始注册表值 = " +
                    (snap.Values.Count == 0 ? "（无）" : JoinKv(snap.Values)));
                sb.AppendLine("  原始 " + "autostart.txt = " + (snap.HadMarker ? snap.Marker.Trim() : "（无）"));
                sb.AppendLine("  当前命令行 = " + AutoStart.CommandLine);

                bool on = AutoStart.SetEnabled(true);
                bool has = AutoStart.IsRegistered(out string v1);
                check("开启后注册表里有这一条", on && has, "value=" + (v1 ?? "null"));
                check("登记的内容就是本程序的命令行",
                    string.Equals(v1, AutoStart.CommandLine, StringComparison.OrdinalIgnoreCase),
                    v1 ?? "null");

                bool off = AutoStart.SetEnabled(false);
                bool has2 = AutoStart.IsRegistered(out string v2);
                check("关闭后注册表里没有了", off && !has2, "value=" + (v2 ?? "null"));

                // ---- 历史值名早期版本写的是 `VbeComplete`，关自启时必须一起清掉 ----
                AutoStart.RawSet("VbeComplete", AutoStart.CommandLine);
                bool seen = AutoStart.IsRegistered(out string v4) &&
                            string.Equals(v4, AutoStart.CommandLine, StringComparison.OrdinalIgnoreCase);
                check("能认出历史值名 VbeComplete 那一条", seen, "value=" + (v4 ?? "null"));

                bool offLegacy = AutoStart.SetEnabled(false);
                bool legacyGone = !AutoStart.IsRegistered(out string v5);
                check("关闭自启时把历史值名一起清掉", offLegacy && legacyGone, "value=" + (v5 ?? "null"));
            }
            finally
            {
                AutoStart.Restore(snap);
            }

            // 还原判定：主名 + 历史值名全都要和快照一致（只看主名会漏掉旧值名那条）
            var after = AutoStart.Capture();
            bool same = after.Values.Count == snap.Values.Count;
            foreach (var kv in snap.Values)
            {
                if (!after.Values.TryGetValue(kv.Key, out string av) ||
                    !string.Equals(av, kv.Value, StringComparison.OrdinalIgnoreCase))
                    same = false;
            }
            check("测完已还原成测试前的状态", same,
                "before=" + (snap.Values.Count == 0 ? "无" : JoinKv(snap.Values)) +
                " / after=" + (after.Values.Count == 0 ? "无" : JoinKv(after.Values)));

            sb.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "startuptest.log"), sb.ToString(), Encoding.UTF8);
            return ok ? 0 : 1;
        }

        /// <summary>把 Run 键快照拼成一行，只用于 --startuptest 的失败详情。</summary>
        private static string JoinKv(Dictionary<string, string> d)
        {
            var sb = new StringBuilder();
            foreach (var kv in d)
            {
                if (sb.Length > 0) sb.Append("; ");
                sb.Append(kv.Key).Append("=").Append(kv.Value);
            }
            return sb.ToString();
        }
    }
}
