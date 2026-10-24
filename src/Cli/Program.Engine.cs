// 候选规则的真机回归：--enginetest。
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
        /// 候选判定的**真机**回归：拿真实模块文本 + 人造快照跑 Trigger，
        /// 逐条验证"这个位置到底弹不弹、弹出的第一条是什么"。
        ///
        /// 为什么不直接敲键盘：钩子按设计不理注入键，没法在这里模拟打字。
        ///   但 Trigger 的全部输入就是 (模块文本, 行, 列)，人造快照与真打字
        ///   走的是同一条代码路径 —— 判定逻辑 100% 覆盖，只有"按键有没有送到
        ///   钩子"覆盖不到（那部分由 --hktest 的命令级用例兜）。
        ///
        /// 本命令纯读，一个字符都不写模块。
        /// </summary>
        private static int EngineTest()
        {
            var sb = new StringBuilder();
            bool ok = true;

            // 真机信息只打印、不作判据 —— 没开代码窗时这里是 -1。
            // 不能把整节押在"真机模块里恰好有 student2 / okkk2 / 他们"上，
            //   换个模块就整片红或整片 skip —— 那样的绿说明不了任何事。
            string liveText = null;
            int liveLines = -1;
            try
            {
                var live = new VbeBackend();
                liveText = live.ModuleText();
                liveLines = live.LineCount();
            }
            catch { }

            sb.AppendLine("=== engine test（合成工程，不依赖真机打开的模块）===");
            sb.AppendLine("真机模块（仅供参考）：lines=" + liveLines +
                          " len=" + (liveText == null ? -1 : liveText.Length));

            // 合成工程：一个标准模块（模块1）+ 一个类模块（cls1）。
            // 组件类型要真的填（1=标准 2=类）—— "类成员不进裸名池"就靠它判。
            const string modStd =
                "Option Explicit\n" +
                "\n" +
                "Public 我们 As String\n" +
                "Private okkk2 As String\n" +
                "\n" +
                "Sub test()\n" +
                "    Dim student2 As String\n" +
                "    Dim 他们 As String\n" +
                "    Dim dic As Object, c As cls1\n" +
                "    student = getOne()\n" +
                "    \n" +
                "End Sub\n";
            const string modCls =
                "Option Explicit\n" +
                "\n" +
                "Public s1 As String\n" +
                "Public s2 As String\n";
            // 故意让这个模块里只有 `wsR.Cells(...)` 这种**成员访问**。
            //   若把它拆成"wsR"和"Cells"两个自由标识符，Cells 就会进全工程
            //   的裸名池 —— 用户在**任何**模块打 `cell` 都会冒出 `Cells`。
            const string modMem =
                "Sub W()\n" +
                "    n = wsR.Cells(1, 1).Value\n" +
                "End Sub\n";
            // 无 Option Explicit 的 Worksheet 模块（真机样例）。
            //   `Range` 只在两处出现：`As Range`（类型）与 `Range("d2")`（调用），
            //   都不是"自动声明的变量" —— 打 `ra` 不该冒出 Range。
            const string modNoExp =
                "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\n" +
                "    Set Target = Target.Cells(1, 1)\n" +
                "    If Target.Address = Range(\"d2\").Address Then\n" +
                "        setValidity ActiveCell, \"c\"\n" +
                "    End If\n" +
                "    ra\n" +
                "    n7 = Selection.End(xlUp)\n" +
                "End Sub\n";

            // 无 Option Explicit 的隐式变量是**模块私有**的 —— 另一个标准模块
            //   里写过的 `secretVar` 不该在别处冒出来。模块2 里同时放一行 `sec`，
            //   用来证明"在自己的模块里照常提示"。
            const string modFree = "Sub Q()\n" +
                                   "    secretVar = 1\n" +
                                   "    sec\n" +
                                   "End Sub\n";
            // Type 块里的成员（只能 p.Nickname 访问）不是裸名
            const string modType = "Public Type Person\n" +
                                   "    Nickname As String\n" +
                                   "    Age As Long\n" +
                                   "End Type\n";
            // 真机样例（无 Option Explicit 的 Worksheet 模块）：
            //   `Set conditions = getConditions()` 之后，在 `If isMeet(condi|` 里
            //   打 condi 提示不出 conditions。
            //   关键：下面 `Private Function isMeet(conditions, ws, t, r)` 的**形参**
            //   也叫 conditions —— 按名字一刀切排除的话，形参会把隐式变量挡掉。
            const string modUser =
                "Sub 查询()\n" +
                "    Set ws = Worksheets(\"人员考勤记录\")\n" +
                "    Set t = getTitle(ws, 1, 1)\n" +
                "    Set conditions = getConditions()      ' 已设置的条件\n" +
                "    Set dic = CreateObject(\"scripting.dictionary\")\n" +
                "    For i = 2 To ws.Cells(Cells.Rows.Count, t(\"姓名\")).End(xlUp).Row\n" +
                "        If isMeet(condi\n" +
                "    Next\n" +
                "End Sub\n" +
                "\n" +
                "Private Function getConditions()\n" +
                "    Set dic = CreateObject(\"scripting.dictionary\")\n" +
                "    Set getConditions = dic\n" +
                "End Function\n" +
                "\n" +
                "Private Function isMeet(conditions, ws, t, r)\n" +
                "    If conditions.exists(\"项目名称\") Then\n" +
                "    End If\n" +
                "End Function\n";

            // 第二个案例（无 Option Explicit）：
            //   `Set dic = CreateObject(...)` / `dic.Add s, i` 之后，在
            //   `set getTitle2=di|` 里打 di 提示不出 dic。
            const string modUser2 =
                // 用户真实模块里 getTitle 和 getTitle2 **都有** `Set dic = ...`
                //   —— 同名隐式变量在两个过程里各有一份，只记第一次出现的那个过程，
                //   getTitle2 里就再也看不见它了。
                "Function getTitle(ws, r1)\n" +
                "    Set dic = CreateObject(\"scripting.dictionary\")\n" +
                "    Set getTitle = dic\n" +
                "End Function\n" +
                "\n" +
                "Function getTitle2(ws, r1, r2)\n" +
                "    Set dic = CreateObject(\"scripting.dictionary\")\n" +
                "    lastCol1 = ws.Cells(r1, Cells.Columns.Count).End(xlToLeft).Column\n" +
                "    lastCol2 = ws.Cells(r2, Cells.Columns.Count).End(xlToLeft).Column\n" +
                "    lastCol = Application.WorksheetFunction.Max(lastCol1, lastCol2)\n" +
                "    For i = 1 To lastCol\n" +
                "        s1 = ws.Cells(r1, i).MergeArea.Cells(1, 1)\n" +
                "        s2 = ws.Cells(r2, i).MergeArea.Cells(1, 1)\n" +
                "        s = s1 & s2\n" +
                "        dic.Add s, i\n" +
                "    Next\n" +
                "    set getTitle2=di\n" +
                "End Function\n";

            var synth = new SyntheticBackend()
                .Add("模块1", 1, modStd)
                .Add("cls1", 2, modCls)
                .Add("modMem", 1, modMem)
                .Add("Sheet1", 100, modNoExp)
                .Add("模块2", 1, modFree)
                .Add("modT", 1, modType)
                .Add("模块3", 1, modUser)
                .Add("模块4", 1, modUser2);

            // 说明, 模块, 行, 行文本, 光标显示列, 期望该弹?, 必须有, 不许有
            Action<string, string, int, string, int, bool?, string[], string[]> sprobe =
                (name, mod, line, lineText, col, wantShow, mustHave, mustNot) =>
                {
                    synth.Current = mod;
                    var e = new Engine { Recent = new Recent(), Backend = synth };
                    var sn = new Snapshot
                    {
                        Line = line,
                        Text = lineText,
                        Module = mod,
                        CaretCol = col
                    };
                    e.InvalidatePool();
                    bool shown = e.Trigger(false, sn);
                    var items = shown ? e.Items.ToList() : new List<Candidate>();

                    bool pass = true;
                    if (wantShow.HasValue && shown != wantShow.Value) pass = false;
                    foreach (var want in mustHave ?? Array.Empty<string>())
                        if (!items.Any(x => string.Equals(x.Name, want, StringComparison.OrdinalIgnoreCase)))
                            pass = false;
                    foreach (var bad in mustNot ?? Array.Empty<string>())
                        if (items.Any(x => string.Equals(x.Name, bad, StringComparison.OrdinalIgnoreCase)))
                            pass = false;
                    if (!pass) ok = false;

                    sb.AppendLine((pass ? "  PASS  " : "  FAIL  ") + name);
                    sb.AppendLine("         行[" + (lineText ?? "") + "] 列=" + col + " -> " +
                                  (shown ? string.Join(",", items.Take(10).Select(x => x.Name)) : "(不弹)") +
                                  (mustHave != null && mustHave.Length > 0
                                      ? "   必须有=" + string.Join("/", mustHave) : "") +
                                  (mustNot != null && mustNot.Length > 0
                                      ? "   不许有=" + string.Join("/", mustNot) : ""));
                    if (shown) e.Hide();
                };

            // ---- 1. 声明位置 / `As` 之后不弹 ----
            sprobe("As 后不弹（Dim student2 As Str|）", "模块1", 7,
                "    Dim student2 As String", 24, false, null, null);
            sprobe("裸声明名字位不弹（Private okkk2|）", "模块1", 4,
                "Private okkk2 As String", 14, false, null, null);
            sprobe("裸声明 + As 不弹（Public 我们 As Str|）", "模块1", 3,
                "Public 我们 As String", 19, false, null, null);
            sprobe("裸声明名字位不弹（Public 我|）", "模块1", 3,
                "Public 我们 As String", 10, false, null, null);

            // ---- 2. Option Explicit：未声明的名字不进候选 ----
            // 光标所在行会被"挖掉半截词"，所以搜索词要和模块里的名字错开：
            // 池子里若混进了第 10 行的 `student`，它就会在 stud 这次冒出来。
            sprobe("未声明 student 不进候选（stud|）", "模块1", 11,
                "    stud", 9, null, new[] { "student2" }, new[] { "student" });
            sprobe("未声明 ok 不进候选（ok|）", "模块1", 11,
                "    ok", 7, null, new[] { "okkk2" }, new[] { "ok" });
            sprobe("中文变量名能提示（他|）", "模块1", 11,
                "    他", 7, true, new[] { "他们" }, null);

            // ---- 2b. 无 Option Explicit 时，"只当函数调用过"的名字不算变量 ----
            //   真机样例：Worksheet 模块里打 `ra`，第一条是 Excel 的 `Range` ——
            //   它只是代码里 `Range("d2")` 调用过一次的宿主函数。
            //   内建函数整体不进候选了，`ra` 也不可能再命中 Range。
            //   本模块的过程名仍可能松散命中（Worksheet_SelectionChange），
            //   那是"自己的名字"，允许弹。
            sprobe("无 Option Explicit 时调用名不进变量池（ra| 不冒 Range）", "Sheet1", 6,
                "    ra", 7, true, null, new[] { "Range" });
            // 宿主枚举常量（xlUp / vbYes / mso*）不进池 —— PY 的 _RE_BUILTIN_CONST
            sprobe("宿主枚举常量不进池（xlU| 不冒 xlUp）", "Sheet1", 6,
                "    xlU", 8, null, null, new[] { "xlUp" });
            // 宿主全局对象（Selection / ActiveSheet）不进池
            sprobe("宿主对象名不进池（Sele| 不冒 Selection）", "Sheet1", 6,
                "    Sele", 9, null, null, new[] { "Selection" });

            // ---- 2c. 隐式变量的可见范围 / Type 成员 ----
            sprobe("隐式变量不跨模块泄露（Sheet1 里 sec| 不冒模块2 的 secretVar）", "Sheet1", 6,
                "    sec", 8, null, null, new[] { "secretVar" });
            sprobe("本模块的隐式变量照常提示（模块2 里 sec|）", "模块2", 3,
                "    sec", 8, true, new[] { "secretVar" }, null);
            sprobe("Type 成员不进裸名池（模块1 里 Nic| 不冒 Nickname）", "模块1", 11,
                "    Nic", 8, null, null, new[] { "Nickname" });

            // ---- 2d. 用户案例（Set 出来的隐式变量必须能提示）----
            var dbg = Scope.Analyze("模块3", modUser);
            sb.AppendLine("  [诊断] Decls: " + string.Join(",", dbg.Decls.Select(d => d.Name + "@" + (d.Proc ?? "-"))));
            sb.AppendLine("  [诊断] FreeIdents: " + string.Join(",", dbg.FreeIdents));
            sb.AppendLine("  [诊断] Procs: " + string.Join(",", dbg.Procs.Select(p => p.Name + ":" + p.StartLine + "-" + p.EndLine)));
            sprobe("Set 出来的隐式变量能提示（condi| -> conditions）", "模块3", 7,
                "        If isMeet(condi", 24, true, new[] { "conditions" }, null);
            sprobe("同函数内 Set 出来的 dic 能提示（di| -> dic）", "模块4", 17,
                "    set getTitle2=di", 20, true, new[] { "dic" }, null);

            // ---- 3. A/B 对照：去掉 Option Explicit，同一个探针就该收进来 ----
            //   证明"未声明的不提示"确实是 Option Explicit 在起作用，不是别的原因。
            var noExpStd = modStd.Replace("Option Explicit", "              ");
            var synthNoExp = new SyntheticBackend()
                .Add("模块1", 1, noExpStd)
                .Add("cls1", 2, modCls);
            synthNoExp.Current = "模块1";
            var engNoExp = new Engine { Recent = new Recent(), Backend = synthNoExp };
            bool shown2 = engNoExp.Trigger(false,
                new Snapshot { Line = 11, Text = "    stud", Module = "模块1", CaretCol = 9 });
            bool leaked = shown2 && engNoExp.Items.Any(
                x => string.Equals(x.Name, "student", StringComparison.OrdinalIgnoreCase));
            if (!leaked) ok = false;
            sb.AppendLine((leaked ? "  PASS  " : "  FAIL  ") +
                          "对照：没有 Option Explicit 时 student 会被收进来");
            sb.AppendLine("         -> " +
                          (shown2 ? string.Join(",", engNoExp.Items.Take(8).Select(x => x.Name)) : "(不弹)"));
            if (shown2) engNoExp.Hide();

            // ---- 4. `New` 后面不弹（用户 第 2 条，跟 `As` 一样）----
            sprobe("New 后不弹（dic.Add 1, New |）", "模块1", 11,
                "    dic.Add 1, New ", 21, false, null, null);
            sprobe("New 后半截词也不弹（New cl|）", "模块1", 11,
                "    dic.Add 1, New cl", 22, false, null, null);
            // 这两条必须用"半截词"：`New |` / `As |`（词还没开始）本来就走
            //   "空词不弹"那条路，退掉 AfterTypeKeyword 也照样绿 —— 抓不到问题。
            //   真正会弹出全工程候选的是 `New cl|` / `As Lo|` 这种打了半个字的。
            sprobe("As 后不弹（For i As Lo|，行首不是 Dim）", "模块1", 11,
                "    For i As Lo", 17, false, null, null);

            // ---- 5. 类模块的 Public 成员不许混进别的模块的裸名池 ----
            sprobe("类模块成员不进裸名池（模块1 里打 s|）", "模块1", 11,
                "    s", 6, null, null, new[] { "s1", "s2" });
            // 症状：输入 cell 会提示 cells（泄露）。
            //   modMem 里只有 `wsR.Cells(1, 1).Value` —— 全是成员访问。
            //   在**另一个模块**里打 `cell` 时，候选里不许出现 Cells。
            sprobe("成员名不进裸名池（打 cell 不冒 Cells）", "模块1", 11,
                "    cell", 8, null, null, new[] { "Cells" });
            // wsR 是 modMem 里"用过但没声明"的名字 —— VBA 会把它当成**该模块
            //   私有**的隐式变量。所以在 modMem 自己里必须提示，在别的模块里不该提示
            //   （那里写出来只会新建一个 Empty 变量，运行时报错）。
            sprobe("同一个词继续：在 modMem 自己里打 ws 能提示对象名 wsR", "modMem", 2,
                "    ws", 6, true, new[] { "wsR" }, null);
            sprobe("隐式变量不跨模块：模块1 里打 ws 不冒 wsR", "模块1", 11,
                "    ws", 6, null, null, new[] { "wsR" });
            // 对照：在 cls1 自己里面它就是裸名 —— 必须提示（证明按模块过滤，不是一刀切）
            sprobe("对照：cls1 自己里 s1 照常提示（s|）", "cls1", 4,
                "s", 2, true, new[] { "s1" }, null);
            // ---- 6. 点号之后一整个让给 VBE ----
            // 默认（MEMBER_POPUP=0）一个候选都不出：VBE 在这个位置按真类型弹真成员。
            sprobe("点号后不弹（c.|）", "模块1", 11,
                "    c.", 7, false, null, null);
            sprobe("点号后不弹（c.s|，成员名打到一半）", "模块1", 11,
                "    c.s", 8, false, null, null);
            sprobe(") 之后不弹（Range(\"A1\").Fo|）", "模块1", 11,
                "    Range(\"A1\").Fo", 19, false, null, null);
            // 词要用能匹配到东西的（`.Se` 会命中 Set / Select / …）——
            // 用 `.Size` 这种"恰好一个候选都匹配不上"的词，退掉判据也照样绿，
            // 等于没验的 A/B 已经栽过一次这种"探针空转"）。
            sprobe("With 块点号不弹（    .Se|）", "模块1", 11,
                "    .Se", 9, false, null, null);
            {
                // ---- A/B：MEMBER_POPUP=1 时同一位置必须重新弹（且按类型收窄）----
                // 这一条是"点号后不弹"的**反证**：退掉那条判据它就绿、留着它才红不了；
                // 反过来一旦判据写错（比如把 `c.` 判成非成员位置），这条就红。
                synth.Current = "模块1";
                var eOn = new Engine
                {
                    Recent = new Recent(),
                    Backend = synth,
                    MemberPopup = true
                };
                eOn.InvalidatePool();
                bool shownOn = eOn.Trigger(false,
                    new Snapshot { Line = 11, Text = "    c.", Module = "模块1", CaretCol = 7 });
                bool hasMembers = shownOn && eOn.Items.Any(
                    x => string.Equals(x.Name, "s1", StringComparison.OrdinalIgnoreCase));
                if (!(shownOn && hasMembers)) ok = false;
                sb.AppendLine(((shownOn && hasMembers) ? "  PASS  " : "  FAIL  ") +
                              "A/B：MEMBER_POPUP=1 点号后照弹，按类型收窄到 s1/s2");
                sb.AppendLine("         -> " +
                    (shownOn ? string.Join(",", eOn.Items.Take(8).Select(x => x.Name)) : "(不弹)"));
                if (shownOn) eOn.Hide();
            }

            sb.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "enginetest.log"), sb.ToString(), Encoding.UTF8);
            Console.WriteLine(sb.ToString());
            return ok ? 0 : 1;
        }

        /// <summary>
        /// 完全合成的工程（不碰 COM）：模块名 -> 全文 + 组件类型。
        /// 用来钉"跨模块可见性"这类**与当前打开哪个模块无关**的契约 ——
        /// 真机探针在没开代码窗时一律 skip，等于什么都没验之前
        /// paretest / hktest / enginetest 三条就是这么"绿"着的）。
        /// </summary>
        private sealed class SyntheticBackend : IVbeSource
        {
            private readonly Dictionary<string, string> _texts =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, int> _types =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            public string Current;

            public SyntheticBackend Add(string name, int type, string text)
            {
                _texts[name] = text;
                _types[name] = type;
                return this;
            }

            public Snapshot Snapshot(bool verify = true) => null;

            public string ModuleText() =>
                _texts.TryGetValue(Current ?? "", out var t) ? t : "";

            public List<string> ComponentNames() => new List<string>(_texts.Keys);

            public List<ModuleScope> ProjectScopes(out int signature)
            {
                signature = 1;
                var list = new List<ModuleScope>();
                foreach (var kv in _texts)
                {
                    var ms = Scope.Analyze(kv.Key, kv.Value);
                    ms.ComponentType = _types[kv.Key];
                    list.Add(ms);
                }
                return list;
            }

            public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
                string completion, bool wantParens, bool wantSpace, string enclosingProc) => false;

            public string EnclosingProcName() => "test";
        }
    }
}
