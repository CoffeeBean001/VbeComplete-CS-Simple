// 日志：只在 VBECOMPLETE_LOG=1 时落盘。
// 任何写日志失败都必须被吞掉 —— 日志永远不允许影响功能。
//
// 为什么是"入队 + 后台落盘"而不是直接 AppendAllText：
//   这个函数会被**低级键盘钩子回调**调用（"热键未命中 / 被跳过"那两条诊断）。
//   钩子回调里做文件 IO 是硬伤 —— 磁盘慢、杀软扫描、网络盘一卡，回调超过
//   LowLevelHooksTimeout（默认 300ms），Windows 会直接把我们的钩子摘掉：
//   从那以后全机器的按键都要重新走一遍系统路径，手感发涩，而日志里一个字都没有。
//   入队是纯内存操作（微秒级），真正的 IO 交给后台线程，钩子永远不会被 IO 拖住。

using System;
using System.Collections.Concurrent;
using System.IO;
using System.Text;
using System.Threading;

namespace VbeComplete
{
    internal static class Log
    {
        public static readonly bool Enabled = Settings.Flag("LOG", false);

        /// <summary>单个日志文件的上限，超了就轮转成 .1（再超就覆盖旧的 .1）。</summary>
        private const long MaxBytes = 4L * 1024 * 1024;
        /// <summary>队列上限：后台线程被卡死时宁可丢日志，也不能把内存吃光。</summary>
        private const int MaxQueue = 4000;

        private static readonly ConcurrentQueue<string> _queue =
            new ConcurrentQueue<string>();
        private static readonly object _fileLock = new object();
        private static Thread _writer;
        private static int _dropped;

        public static string Path_
        {
            get
            {
                if (_path == null)
                    _path = System.IO.Path.Combine(Settings.Home, "VbeComplete.log");
                return _path;
            }
        }
        private static string _path;

        public static void Write(string msg)
        {
            if (!Enabled) return;
            Enqueue(Line(msg));
        }

        public static void Write(string fmt, params object[] args)
        {
            if (!Enabled) return;
            string msg;
            try { msg = string.Format(fmt, args); }
            catch { msg = fmt; }                 // 格式串写错也不许影响功能
            Enqueue(Line(msg));
        }

        private static string Line(string msg) =>
            string.Format("{0:HH:mm:ss.fff} [cs] {1}\r\n", DateTime.Now, msg);

        private static void Enqueue(string line)
        {
            if (_queue.Count >= MaxQueue) { Interlocked.Increment(ref _dropped); return; }
            _queue.Enqueue(line);
            EnsureWriter();
        }

        private static void EnsureWriter()
        {
            if (_writer != null) return;
            lock (_fileLock)
            {
                if (_writer != null) return;
                var t = new Thread(DrainLoop) { IsBackground = true, Name = "log-writer" };
                _writer = t;
                t.Start();
            }
        }

        private static void DrainLoop()
        {
            var sb = new StringBuilder(4096);
            while (true)
            {
                Thread.Sleep(250);
                try { Drain(sb); } catch { }     // 落盘失败：等下一轮，绝不往外抛
            }
        }

        /// <summary>把当前排队的日志一次性写下去（退出前调它，别把最后几行丢了）。</summary>
        public static void Flush()
        {
            if (!Enabled) return;
            try { Drain(new StringBuilder(4096)); } catch { }
        }

        private static void Drain(StringBuilder sb)
        {
            sb.Clear();
            string s;
            while (_queue.TryDequeue(out s))
            {
                sb.Append(s);
                if (sb.Length >= 64 * 1024) break;   // 一次别写太大，剩下的下一轮
            }
            int dropped = Interlocked.Exchange(ref _dropped, 0);
            if (dropped > 0) sb.Append("... 队列满，丢弃 " + dropped + " 条日志\r\n");
            if (sb.Length > 0) Append(Path_, sb.ToString(), MaxBytes);
        }

        private static void Append(string path, string text, long maxBytes)
        {
            try
            {
                lock (_fileLock)
                {
                    var fi = new FileInfo(path);
                    if (fi.Exists && fi.Length > maxBytes) Rotate(path);
                    File.AppendAllText(path, text, Encoding.UTF8);
                }
            }
            catch { }
        }

        private static void Rotate(string path)
        {
            try
            {
                string bak = path + ".1";
                if (File.Exists(bak)) File.Delete(bak);
                File.Move(path, bak);
            }
            catch { }
        }

        /// <summary>
        /// 崩溃记录：**不受 LOG 开关控制，也不走队列**，立即落盘。
        /// 原因很直接 —— WinExe 没有控制台，异常被 CLR 吞掉时用户只看到
        /// "双击没反应"，而我们手里一点线索都没有。崩溃恰恰是没开日志时最需要留下的东西，
        /// 而且下一秒进程可能就没了，排队等于丢掉。
        /// </summary>
        public static void Crash(string where, Exception e)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendFormat("==== {0:yyyy-MM-dd HH:mm:ss.fff} {1} ====\r\n", DateTime.Now, where);
                for (var cur = e; cur != null; cur = cur.InnerException)
                {
                    sb.AppendFormat("{0}: {1}\r\n{2}\r\n",
                        cur.GetType().FullName, cur.Message, cur.StackTrace);
                }
                Append(System.IO.Path.Combine(Settings.Home, "crash.log"),
                    sb.ToString(), 1024 * 1024);
            }
            catch { }
        }
    }
}
