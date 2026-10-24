// 让托盘图标"常驻显示"（而不是被 Windows 11 折叠进 ^ 溢出区）。
//
// 为什么需要这一步：Win11 默认把所有新图标的默认状态设成"隐藏"，只留一个 ^。
//   程序自己没法让图标跳到外面（那是用户的显示设置），但 Windows 会把每个托盘
//   图标的状态记在：
//       HKCU\Control Panel\NotifyIconSettings\<一个数字 id>\
//           ExecutablePath  = 哪个 exe
//           IsPromoted      = 1 表示"始终在任务栏显示"
//   我们在启动后找到**自己那条**（按 ExecutablePath 精确比对，绝不碰别人的），
//   把 IsPromoted 置 1。任务栏在下一次读取时生效 —— 实测"重启资源管理器/重新登录"
//   之后就能看到；当前这一次会话里它仍然待在 ^ 里（我们不会去重启别人的 explorer）。
//
// 这条记录是 Windows 在图标**第一次出现时**创建的。所以第一次运行时可能还找不到，
//   下一次启动就找到了 —— 代码两种情况都容得下，找不到就安静返回。
//
// 只改 IsPromoted 一个值，只改 ExecutablePath 等于本程序的子键。可逆、无副作用。

using System;
using Microsoft.Win32;

namespace VbeComplete
{
    internal static class TrayPin
    {
        private const string Root = @"Control Panel\NotifyIconSettings";

        /// <summary>
        /// 尽力把自己的托盘图标设为"始终显示"。只写日志，不抛异常。
        /// </summary>
        public static void EnsureVisible()
        {
            string exe = AutoStart.ExePath();
            if (string.IsNullOrEmpty(exe)) return;

            int seen = 0, changed = 0;
            try
            {
                using (var root = Registry.CurrentUser.OpenSubKey(Root, true))
                {
                    if (root == null) return;          // 老系统 / 精简版没有这个键
                    foreach (string id in root.GetSubKeyNames())
                    {
                        try
                        {
                            using (var k = root.OpenSubKey(id, true))
                            {
                                if (k == null) continue;
                                string p = k.GetValue("ExecutablePath") as string;
                                if (string.IsNullOrEmpty(p)) continue;
                                if (!string.Equals(p, exe, StringComparison.OrdinalIgnoreCase)) continue;

                                seen++;
                                object raw = k.GetValue("IsPromoted");
                                int cur = raw is int i ? i : (raw is long l ? (int)l : 0);
                                if (cur != 1)
                                {
                                    k.SetValue("IsPromoted", 1, RegistryValueKind.DWord);
                                    changed++;
                                }
                            }
                        }
                        catch (Exception e) { Log.Write("托盘常驻：处理 {0} 失败：{1}", id, e.Message); }
                    }
                }
                if (changed > 0)
                    Log.Write("托盘常驻：已把本程序图标设为始终显示（{0} 条，重启资源管理器/重新登录后生效）",
                        changed);
                else if (seen > 0)
                    Log.Write("托盘常驻：已是最新（找到 {0} 条，无需改动）", seen);
                else
                    Log.Write("托盘常驻：还没找到本程序的托盘记录（首次运行属正常，下次启动会再试）");
            }
            catch (Exception e) { Log.Write("托盘常驻：写入失败（不影响使用）：{0}", e.Message); }
        }
    }
}
