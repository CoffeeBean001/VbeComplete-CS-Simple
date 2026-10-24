// 开机自动启动：往 HKCU\Software\Microsoft\Windows\CurrentVersion\Run 里登记一条。
//
// 为什么用 HKCU 的 Run 键，而不是启动文件夹 / 计划任务：
//   1) 不需要管理员权限（UAC 一次都不弹），装在哪都能开；
//   2) 只有当前用户生效 —— 这是"我的编辑器插件"，不该影响别的账户；
//   3) 卸载/关掉就是删一个值，随时可逆，不碰系统计划任务库。
//
// "用户的选择"与"注册表当前状态"分开存：
//   注册表是**结果**（可以被别的优化软件删掉、被清理工具改掉），
//   `autostart.txt`（内容就一个 1/0）是**用户的意图**。
//   每次启动都按意图把注册表"修一遍"：
//     · 用户勾了 → 值不在 / 指向旧路径（exe 挪过地方）→ 重写
//     · 用户没勾 → 值还在（比如上次被强行写回）→ 删掉
//   这样换目录、被清理软件动过，都能自愈，不需要用户重新勾一次。
//
// 默认是**开**（用户明确要求"开机自动启动"）；第一次运行没 marker 时按
//   settings.ini 的 AUTO_START 走（默认 1）。想关：托盘菜单取消勾选，
//   或者 settings.ini 里写 AUTO_START = 0 并把 autostart.txt 删掉。
//
// 只认“自己这一个值名”是不够的：值名可能随版本变化（`VbeComplete` /
//   `VbeComplete-CS`），只删当前这一个，旧值永远删不掉 —— 表现为
//   “autostart.txt 已经改成 0 了，开机还是被拉起来”。
//   所以维护一张“已知值名”表，并且删除前一律用 IsOurCommand 判断
//   “这条值指向的到底是不是我们这份 exe”，绝不误删别家的同名键。

using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;

namespace VbeComplete
{
    internal static class AutoStart
    {
        private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

        /// <summary>当前使用的值名。</summary>
        private const string ValueName = "VbeComplete-CS";

        /// <summary>可能存在的值名 —— 清理 / 识别时一并算作"我们自己"。</summary>
        private static readonly string[] LegacyNames = { "VbeComplete" };

        private const string StateFile = "autostart.txt";

        private static string StatePath => Path.Combine(Settings.Home, StateFile);

        /// <summary>主名 + 历史别名，按"先主后旧"排列。</summary>
        private static string[] AllNames()
        {
            var list = new List<string>(1 + LegacyNames.Length);
            list.Add(ValueName);
            list.AddRange(LegacyNames);
            return list.ToArray();
        }

        /// <summary>要登记的完整命令行（带引号，防路径里有空格）。</summary>
        public static string CommandLine
        {
            get
            {
                string exe = ExePath();
                if (string.IsNullOrEmpty(exe)) return null;
                return "\"" + exe + "\"";
            }
        }

        /// <summary>
        /// 当前进程的真实 exe 路径。单文件发布时 `ProcessPath` 就是 `dist\VbeComplete.exe`；
        /// 极端情况下取不到，退回 `Application.ExecutablePath`（apphost 路径）。
        /// </summary>
        public static string ExePath()
        {
            try
            {
                string p = Environment.ProcessPath;
                if (!string.IsNullOrEmpty(p) &&
                    p.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    // `dotnet run` / `dotnet xxx.dll` 启动时 ProcessPath 是 dotnet.exe。
                    //   把那个登记进 Run 键，开机只会跑一个什么参数都没有的 dotnet
                    //   （窗口一闪就没了）。这种情况退回 apphost 路径。
                    string stem = Path.GetFileNameWithoutExtension(p);
                    if (!string.Equals(stem, "dotnet", StringComparison.OrdinalIgnoreCase))
                        return p;
                }
            }
            catch { }
            try { return System.Windows.Forms.Application.ExecutablePath; } catch { }
            return null;
        }

        /// <summary>从注册表里的命令行取出 exe 路径（能处理 `"C:\a b\x.exe" -arg`）。</summary>
        private static string ExeIn(string cmdline)
        {
            if (string.IsNullOrWhiteSpace(cmdline)) return "";
            string s = cmdline.Trim();
            if (s[0] == '"')
            {
                int q = s.IndexOf('"', 1);
                if (q > 0) return s.Substring(1, q - 1).Trim();
            }
            // 没引号：退到第一个 ".exe" 为止（路径里可能有空格，不能简单按空格切）
            int e = s.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
            if (e > 0) return s.Substring(0, e + 4).Trim();
            return s;
        }

        /// <summary>
        /// 这条注册表值是不是"我们这份 exe"的。判据分两级：
        ///   1) 同目录、同名 → 一定是；
        ///   2) 不同目录但那个文件已经不在了 → 是登记后挪走/删掉的失效项，也算。
        /// 只看文件名不够：别家可能也有个 VbeComplete.exe，那种文件还在的不许动。
        /// </summary>
        internal static bool IsOurCommand(string value)
        {
            string exe = ExeIn(value);
            if (string.IsNullOrEmpty(exe)) return false;
            string me = ExePath();
            if (string.IsNullOrEmpty(me)) return false;
            if (!string.Equals(Path.GetFileName(exe), Path.GetFileName(me),
                               StringComparison.OrdinalIgnoreCase))
                return false;
            if (SameDir(exe, me)) return true;
            try { return !File.Exists(exe); } catch { return false; }
        }

        /// <summary>Run 键里所有"属于我们"的项：(值名, 命令行)。</summary>
        internal static List<KeyValuePair<string, string>> Ours()
        {
            var list = new List<KeyValuePair<string, string>>();
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RunKey, false))
                {
                    if (k == null) return list;
                    foreach (var name in AllNames())
                    {
                        string v = k.GetValue(name) as string;
                        if (string.IsNullOrWhiteSpace(v)) continue;
                        if (!IsOurCommand(v)) continue;
                        list.Add(new KeyValuePair<string, string>(name, v));
                    }
                }
            }
            catch (Exception e)
            {
                Log.Write("读开机自启注册表失败：{0}", e.Message);
            }
            return list;
        }

        /// <summary>注册表里现在有没有"我们这一份"（主名优先，附带把值读出来）。</summary>
        public static bool IsRegistered(out string value)
        {
            value = null;
            var ours = Ours();
            if (ours.Count == 0) return false;
            foreach (var kv in ours)
            {
                if (kv.Key == ValueName) { value = kv.Value; return true; }
            }
            value = ours[0].Value;
            return true;
        }

        /// <summary>用户意图：读 marker，没有就按配置（默认开）。</summary>
        public static bool Desired
        {
            get
            {
                try
                {
                    if (File.Exists(StatePath))
                    {
                        string s = File.ReadAllText(StatePath).Trim();
                        if (s == "0") return false;
                        if (s == "1") return true;
                    }
                }
                catch { }
                return Settings.Flag("AUTO_START", true);
            }
        }

        /// <summary>用户从托盘菜单切换：先记住意图，再立刻落到注册表。</summary>
        public static bool SetEnabled(bool on)
        {
            try
            {
                File.WriteAllText(StatePath, on ? "1" : "0");
            }
            catch (Exception e) { Log.Write("写 {0} 失败（意图没记住，本次仍然生效）：{1}", StateFile, e.Message); }
            return Apply(on);
        }

        /// <summary>
        /// 启动时调用：按意图把注册表修正到"该有的样子"。
        /// 已一致就一个字节都不写（省得每次开机都碰注册表）。
        /// </summary>
        public static void Sync()
        {
            var ours = Ours();

            if (!Desired)
            {
                // 没勾：值还在就删掉（别的工具写回来的、早期版本留下的旧值名，一并清掉）
                if (ours.Count > 0)
                {
                    var names = new List<string>();
                    foreach (var kv in ours) names.Add(kv.Key);
                    if (Apply(false))
                        Log.Write("开机自启：已按你的设置关闭并清理注册表（删了 {0} 条：{1}）",
                                  ours.Count, string.Join(", ", names.ToArray()));
                }
                return;
            }

            string cmd = CommandLine;
            if (cmd == null) { Log.Write("开机自启：取不到 exe 路径，跳过"); return; }

            // 已经是"恰好一条、主名、内容对得上"就不用动
            if (ours.Count == 1 &&
                ours[0].Key == ValueName &&
                string.Equals(ours[0].Value, cmd, StringComparison.OrdinalIgnoreCase))
                return;

            // "路径不一样就直接改写"是危险的：任何一次从 bin\Release、探针目录或
            //   临时解压目录运行，都会把用户装好的**正式版**登记顶掉（真机上发生过
            //   两次，表现为开机后跑的是调试产物、甚至跑不起来）。
            //   规则是：主名那条指向的文件还在、且和当前 exe 不在同一个目录 ⇒ 只记日志，
            //   绝不覆盖。只有"旧值已经失效"或"同一个目录里换了文件名"才自动修。
            string mainVal = null;
            foreach (var kv in ours)
            {
                if (kv.Key == ValueName) { mainVal = kv.Value; break; }
            }
            if (mainVal != null)
            {
                string oldExe = ExeIn(mainVal);
                bool oldAlive = !string.IsNullOrEmpty(oldExe) && File.Exists(oldExe);
                if (oldAlive && !SameDir(oldExe, ExeIn(cmd)))
                {
                    Log.Write("开机自启：注册表指向另一个目录的 {0}（文件还在），本次不覆盖；" +
                              "要换成本目录的 exe，请在托盘菜单里关一次再打开", oldExe);
                    return;
                }
            }

            if (Apply(true))
                Log.Write("开机自启：已登记 {0}（原值 {1}）", cmd, mainVal ?? "无");
        }

        private static bool SameDir(string a, string b)
        {
            try
            {
                string da = Path.GetDirectoryName(Path.GetFullPath(a));
                string db = Path.GetDirectoryName(Path.GetFullPath(b));
                return string.Equals(da, db, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        /// <summary>写 / 删注册表。on=true 时顺手把历史值名归一（只留主名一条）。</summary>
        private static bool Apply(bool on)
        {
            string cmd = CommandLine;
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RunKey, true))
                {
                    if (k == null) { Log.Write("开机自启：打不开 Run 键"); return false; }
                    if (on)
                    {
                        if (cmd == null) return false;
                        k.SetValue(ValueName, cmd, RegistryValueKind.String);
                        // 归一：指向我们这份 exe 的旧值名删掉，免得开机被拉起两次
                        foreach (var n in LegacyNames)
                        {
                            string v = k.GetValue(n) as string;
                            if (string.IsNullOrWhiteSpace(v)) continue;
                            if (IsOurCommand(v))
                            {
                                k.DeleteValue(n, false);
                                Log.Write("开机自启：已把旧值名 {0}（{1}）归一到 {2}", n, v, ValueName);
                            }
                        }
                        return true;
                    }

                    // 关闭：所有"我们的"值名一起删；同名但指向别处有效程序的键不动
                    foreach (var n in AllNames())
                    {
                        string v = k.GetValue(n) as string;
                        if (v == null) continue;
                        if (IsOurCommand(v))
                        {
                            k.DeleteValue(n, false);
                            Log.Write("开机自启：已删除注册表项 {0} = {1}", n, v);
                        }
                        else
                        {
                            Log.Write("开机自启：{0} 指向别处的同名程序（{1}），不是我们的，不动", n, v);
                        }
                    }
                    return true;
                }
            }
            catch (Exception e)
            {
                Log.Write("开机自启：写注册表失败（可能被安全软件拦了）：{0}", e.Message);
                return false;
            }
        }

        /// <summary>给日志 / 托盘提示用的一句话状态。</summary>
        public static string Describe()
        {
            var ours = Ours();
            string extra = "";
            foreach (var kv in ours)
            {
                if (kv.Key != ValueName) extra += kv.Key + " ";
            }
            extra = extra.Trim();

            if (Desired)
            {
                if (ours.Count == 0)
                    return "开机自启：应开启，但注册表里没有（可能被清理软件删了，下次启动会自愈）";
                if (ours.Count > 1)
                    return "开机自启：已开启，另有旧值名 " + extra + "（下次启动会归一成一条）";
                return "开机自启：已开启";
            }
            if (ours.Count == 0) return "开机自启：已关闭";
            return "开机自启：应关闭，但注册表里还有 " + ours.Count + " 条（含旧值名 " + extra +
                   "），下次启动会一起清掉";
        }

        // ------------------------------------------------------------------
        // 自检用：先拍一张快照，测完原样还原 —— --startuptest 绝不许留下痕迹
        // ------------------------------------------------------------------

        internal sealed class State
        {
            /// <summary>值名 → 原始内容（主名与历史别名都拍下来）。</summary>
            internal readonly Dictionary<string, string> Values =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            internal bool HadMarker;
            internal string Marker;


            internal string Value
            {
                get { string v; return Values.TryGetValue(ValueName, out v) ? v : null; }
            }
        }

        internal static State Capture()
        {
            var st = new State();
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RunKey, false))
                {
                    if (k != null)
                    {
                        foreach (var n in AllNames())
                        {
                            object raw = k.GetValue(n);
                            if (raw != null) st.Values[n] = raw as string ?? "";
                        }
                    }
                }
            }
            catch { }
            try
            {
                if (File.Exists(StatePath))
                {
                    st.HadMarker = true;
                    st.Marker = File.ReadAllText(StatePath);
                }
            }
            catch { }
            return st;
        }

        internal static void Restore(State st)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RunKey, true))
                {
                    if (k != null)
                    {
                        foreach (var n in AllNames())
                        {
                            if (st.Values.TryGetValue(n, out string v))
                            {
                                k.SetValue(n, v ?? "", RegistryValueKind.String);
                            }
                            else
                            {
                                string cur = k.GetValue(n) as string;
                                if (cur != null && IsOurCommand(cur)) k.DeleteValue(n, false);
                            }
                        }
                    }
                }
            }
            catch (Exception e) { Log.Write("还原开机自启注册表失败：{0}", e.Message); }
            try
            {
                if (st.HadMarker) File.WriteAllText(StatePath, st.Marker ?? "");
                else if (File.Exists(StatePath)) File.Delete(StatePath);
            }
            catch (Exception e) { Log.Write("还原 {0} 失败：{1}", StateFile, e.Message); }
        }

        /// <summary>自检专用：直接写某个值名的原始内容（不做"是不是我们的"判断）。</summary>
        internal static void RawSet(string name, string value)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RunKey, true))
                {
                    if (k != null) k.SetValue(name, value ?? "", RegistryValueKind.String);
                }
            }
            catch (Exception e) { Log.Write("自检写注册表 {0} 失败：{1}", name, e.Message); }
        }
    }
}
