// 片段库诊断：--snipetest（临时库，不碰真库）。
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
        /// 片段库自检（不需要 Excel、不需要桌面）：在**临时库**上跑一遍增删改查，
        /// 顺便报告真实库的路径与条数。
        /// 刻意不碰用户的 snippets-cs.db —— 自检绝不允许改坏用户攒下来的片段。
        /// </summary>
        private static int SnippetTest()
        {
            var sb = new StringBuilder();
            bool ok = true;
            string tmp = Path.Combine(Path.GetTempPath(),
                "vbecomplete-snipetest-" + Guid.NewGuid().ToString("N") + ".db");

            sb.AppendLine("=== snippet store test ===");
            sb.AppendLine("winsqlite3 路径 = " + Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System), "winsqlite3.dll"));
            sb.AppendLine("临时库 = " + tmp);
            sb.AppendLine("真实库 = " + SnippetStore.DbPath() +
                          "（存在=" + File.Exists(SnippetStore.DbPath()) + "）");

            Action<string, bool, string> check = (name, pass, detail) =>
            {
                if (!pass) ok = false;
                sb.AppendLine((pass ? "  PASS  " : "  FAIL  ") + name + (detail.Length > 0 ? "   " + detail : ""));
            };

            // 走环境变量换库（Settings.Raw 每次都现读，所以这里设了就生效）
            string oldEnv = Environment.GetEnvironmentVariable("VBECOMPLETE_SNIPPET_DB");
            Environment.SetEnvironmentVariable("VBECOMPLETE_SNIPPET_DB", tmp);
            try
            {
                using (var st = new SnippetStore())
                {
                    check("打开临时库", st.Count() == 0, "count=" + st.Count() + " err=" + st.LastError());

                    check("新增", st.Add("Hello", "MsgBox \"hi\""), "err=" + st.LastError());
                    check("新增第二条", st.Add("Loop", "For i = 1 To 3\r\nNext"), "");
                    check("重名新增返回 false", !st.Add("Hello", "x"), "");
                    check("读回来一致", st.Get("Hello") == "MsgBox \"hi\"", "got=[" + st.Get("Hello") + "]");
                    check("Exists", st.Exists("Loop") && !st.Exists("nope"), "");
                    check("修改代码", st.Update("Hello", "MsgBox \"bye\"") &&
                          st.Get("Hello") == "MsgBox \"bye\"", "");

                    var all = st.All();
                    check("全部按标题排序", all.Count == 2 && all[0].Key == "Hello" && all[1].Key == "Loop",
                        "count=" + all.Count);

                    check("改名", st.Rename("Hello", "Hi") && st.Get("Hi") != null && !st.Exists("Hello"), "");
                    check("改名撞车返回 false", !st.Rename("Hi", "Loop"), "");

                    check("中文标题与多行代码",
                        st.Add("中文片段", "Sub 测试()\r\n    Debug.Print \"中文\"\r\nEnd Sub") &&
                        (st.Get("中文片段") ?? "").Contains("中文"), "");

                    check("删除", st.Delete("Hi") && !st.Exists("Hi"), "");
                    check("删除不存在返回 false", !st.Delete("Hi"), "");
                }

                // 重新打开一次：确认是真的落盘了（不是内存里的假象）
                using (var st2 = new SnippetStore())
                {
                    check("重新打开后数据还在", st2.Count() == 2, "count=" + st2.Count());
                }
            }
            catch (Exception e)
            {
                ok = false;
                sb.AppendLine("EXCEPTION " + e);
            }
            finally
            {
                Environment.SetEnvironmentVariable("VBECOMPLETE_SNIPPET_DB", oldEnv);
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            }

            sb.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "snipetest.log"), sb.ToString(), Encoding.UTF8);
            return ok ? 0 : 1;
        }
    }
}
