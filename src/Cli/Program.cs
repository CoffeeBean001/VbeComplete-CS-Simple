// 入口：单实例、命令行分发、正常启动。
//
// 没有主窗口：进程活着就在托盘里等着，键盘钩子在后台跑。

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
        private static Mutex _mutex;

        /// <summary>已经有一份在跑时，用它叫那一份把托盘图标提示出来。</summary>
        public const string ShowEventName = @"Local\VbeComplete_show";

        [STAThread]
        private static int Main(string[] args)
        {
            // 崩溃必须留下痕迹。WinExe 没有控制台，CLR 把异常吞掉之后用户只看到
            // "窗口闪退、项目没启动"，我们手里连一行线索都没有。
            AppDomain.CurrentDomain.UnhandledException +=
                (s, e) => Log.Crash("AppDomain.UnhandledException", e.ExceptionObject as Exception);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException +=
                (s, e) => Log.Crash("Application.ThreadException", e.Exception);

            try { return RunMain(args); }
            catch (Exception e)
            {
                Log.Crash("Main", e);
                try
                {
                    MessageBox.Show(
                        "启动失败：" + e.Message + "\r\n\r\n详情见程序目录下的 crash.log",
                        "VbeComplete", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
                return 1;
            }
        }

        private static int RunMain(string[] args)
        {
            Settings.Load();
            // 诊断命令把中文打到控制台时，默认代码页（936）会吐一串乱码 ——
            // 这些输出同时会落盘成 UTF-8 日志，能对上才不会看错。
            try { Console.OutputEncoding = Encoding.UTF8; } catch { }

            // 只接受**一个**开关。同时给两个（`--selftest --enterprobe`）若只跑第一个，
            //   `-enterprobe`（单横线）会直接当成"正常启动"—— 两种都会让人以为
            //   "检查跑过了"，其实什么都没跑。多余 / 认不出的参数一律报错退出。
            // `--pooldump <模块名>` 是唯一允许带参数的开关（离线诊断用）。
            if (args.Length > 1 && !(args.Length == 2 && args[0] == "--pooldump"))
            {
                Console.WriteLine("最多只接受一个开关，这次收到 " + args.Length + " 个：" +
                    string.Join(" ", args));
                return 2;
            }

            if (args.Length >= 1)
            {
                switch (args[0])
                {
                    case "--selftest":    return SelfTest.Run();
                    case "--probe":       return Probe();
                    case "--yieldprobe":  return YieldProbe();
                    case "--writetest":   return WriteTest();
                    case "--uitest":      return UiTest();
                    case "--carettest":   return CaretTest();
                    case "--hktest":      return HkTest();
                    case "--enterprobe":  return EnterProbe();
                    case "--snipetest":   return SnippetTest();
                    case "--dump":        return DumpModule();
                    // 诊断用 —— 把当前模块的解析结果 + 光标处的候选全打出来，
                    //   池子里"该有的名字没有"时让用户跑一次，比来回猜快得多。
                    case "--pooldump":    return PoolDump(args.Length > 1 ? args[1] : null);
                    case "--paretest":    return ParenTest();
                    case "--enginetest":  return EngineTest();
                    case "--startuptest": return StartupTest();
                }
            }

            // 认不出的参数不能顺手当成"正常启动"：那会静默起一个托盘实例，
            //   让人以为"检查跑过了"（其实什么都没跑）。打一行用法就退出。
            //   单横线（`-enterprobe`）也算认不出 —— 判据只认 `--` 的话，
            //   敲错一个横线就变成"悄悄启动一个常驻进程"。
            if (args.Length > 0)
            {
                const string usage =
                    "用法：VbeComplete.exe [开关]\r\n" +
                    "  （不带参数 = 正常启动）\r\n" +
                    "  --selftest    纯逻辑自检（不需要 Excel）\r\n" +
                    "  --snipetest   片段库读写检查\r\n" +
                    "  --uitest      候选窗 / 片段窗离屏渲染出 PNG\r\n" +
                    "  --probe       打印当前 VBE 状态（只读）\r\n" +
                    "  --yieldprobe  只采样一次「VBE 自带列表是否可见」（只读）\r\n" +
                    "  --dump        把当前代码窗全文写到 moddump.txt（只读）\r\n" +
                    "  --pooldump    把候选池解析结果 + 光标处候选写到 pooldump.txt（只读）\r\n" +
                    "  --enginetest  用真机模块文本跑候选判定（只读）\r\n" +
                    "  --paretest    真机验证函数补括号（改完即还原）\r\n" +
                    "  --carettest   真机验证中文名光标落点（改完即还原）\r\n" +
                    "  --writetest   真机验证写入链路（改完即还原）\r\n" +
                    "  --hktest      真机验证行操作快捷键（改完即还原）\r\n" +
                    "  --enterprobe  真机验证「回车自动补结构」（走真键盘钩子，改完即还原）\r\n" +
                    "  --startuptest 图标 + 开机自启的登记/回滚（会短暂改注册表，随后还原）\r\n";
                Console.WriteLine("未知开关：" + args[0]);
                Console.WriteLine(usage);
                try
                {
                    File.WriteAllText(Path.Combine(Settings.Home, "usage.log"), usage, Encoding.UTF8);
                }
                catch { }
                return 2;
            }

            // 单实例：同名互斥体已经存在就说明有一份在跑。
            bool created;
            _mutex = new Mutex(true, "VbeComplete", out created);
            if (!created)
            {
                // 只写日志是不够的 —— 用户看到的就是"双击了，什么都没发生"。
                // 叫已经在跑的那一份弹个气泡，把托盘图标指出来。
                SignalExisting();
                Log.Write("已有一份在运行，本次退出（已提示已有实例）");
                return 0;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var app = new App();
            app.Start();

            // Ctrl+C / 关控制台 时收干净。必须回 UI 线程：Shutdown 要释放 NotifyIcon
            // 和窗口，从控制台线程直接 Dispose 是跨线程操作，托盘图标会残留。
            Console.CancelKeyPress += (s, e) => { e.Cancel = true; app.RunOnUi(app.Shutdown); };

            Application.Run(app);
            app.Shutdown();
            try { _mutex.ReleaseMutex(); } catch { }
            return 0;
        }

        private static void SignalExisting()
        {
            EventWaitHandle h;
            if (EventWaitHandle.TryOpenExisting(ShowEventName, out h))
            {
                try { h.Set(); } catch { }
                h.Dispose();
            }
        }
    }
}
