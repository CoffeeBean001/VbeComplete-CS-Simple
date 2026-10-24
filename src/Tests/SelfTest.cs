// 自检：把"补全准不准"的全部判据跑一遍，结果写进 selftest.log。
//
// 用文件而不是 Console：程序是 WinExe（没有控制台），而自检必须在没有桌面的
// 环境下也能跑。

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VbeComplete
{
    internal static class SelfTest
    {
        private static int _pass, _fail;
        private static readonly StringBuilder _sb = new StringBuilder();

        private static void Ok(string name, bool cond, string detail = "")
        {
            if (cond) { _pass++; }
            else
            {
                _fail++;
                _sb.AppendLine("FAIL  " + name + (detail.Length > 0 ? "  <- " + detail : ""));
            }
        }

        private static void Eq<T>(string name, T got, T want)
        {
            bool ok = EqualityComparer<T>.Default.Equals(got, want);
            Ok(name, ok, ok ? "" : "got=" + got + " want=" + want);
        }

        public static int Run()
        {
            _sb.AppendLine("=== VbeComplete C# selftest ===");
            _sb.AppendLine("kernel " + Environment.Version + "  64bit=" + Environment.Is64BitProcess);
            _sb.AppendLine("home   " + Settings.Home);
            _sb.AppendLine();

            try
            {
                TestRules();
                TestParser();
                TestEngine();
                TestPoolContents();
                TestSingleCharName();
                TestFuzzy();
                TestScope();
                TestChineseAndExplicit();
                TestGuards();
                TestFreeIdents();
                TestMember();
                TestEditing();
                TestContinuation();
                TestAutoTrigger();
                TestAutoStart();
                TestSnippets();
            }
            catch (Exception e)
            {
                _fail++;
                _sb.AppendLine("EXCEPTION " + e);
            }

            _sb.AppendLine();
            _sb.AppendLine($"PASS {_pass}   FAIL {_fail}");
            string path = Path.Combine(Settings.Home, "selftest.log");
            try { File.WriteAllText(path, _sb.ToString(), new UTF8Encoding(false)); } catch { }
            return _fail == 0 ? 0 : 1;
        }

        // ------------------------------------------------------------------
        private static void TestRules()
        {
            int ws;
            Eq("word-before/末尾", Rules.WordBefore("Dim abc", 8, out ws), "abc");
            Eq("word-before/起始列", ws, 5);
            Eq("word-before/词中间", Rules.WordBefore("abcdef", 4, out ws), "abc");
            Eq("word-before/无词", Rules.WordBefore("x = ", 5, out ws), "");
            Eq("word-before/空行", Rules.WordBefore("", 1, out ws), "");

            Eq("replace-word", Rules.ReplaceWord("Dim abc", 5, 8, "abcd"), "Dim abcd");
            Eq("replace-word/带后缀", Rules.ReplaceWord("MsgBox \"x\",vbExclamation:Exit Su", 31, 33, "Sub"),
                "MsgBox \"x\",vbExclamation:Exit Sub");

            Eq("dispcol/纯ascii", Rules.DispCol("    abc", 7), 8);
            Eq("dispcol/tab", Rules.DispCol("\tabc", 3), 7);
            Eq("dispcol/宽字符", Rules.DispCol("中文a", 3), 6);

            // Tab 宽度可配置（默认 4，settings.ini 的 TAB_WIDTH 可改）
            int oldTab = Rules.TabWidth;
            try
            {
                Rules.TabWidth = 8;
                Eq("dispcol/tab 宽度可配置（8）", Rules.DispCol("\tabc", 3), 11);
                Eq("char-col/tab 宽度可配置（8）", Rules.CharColFromDisp("\tabc", 11), 4);
            }
            finally { Rules.TabWidth = oldTab; }

            Eq("find-end/同位置", Rules.FindCompletionEnd("    abc = 1", "abc", 7, 7), 7);
            Eq("find-end/被重排后", Rules.FindCompletionEnd("MsgBox \"已\", vbEx: Exit Sub", "Sub", 24, 22), 26);
            // 整词匹配 —— 右边还出现同名子串时，靠 ±12 窗口 + LastIndexOf 会取到
            //   右边那个，把光标甩到词尾之后。
            Eq("find-end/右边同名整词", Rules.FindCompletionEnd("Split=Split(x)", "Split", 0, 0), 5);
            // `=` 补空格后词起点右移（"student=spl" 就是这个）。
            Eq("find-end/等号补空格", Rules.FindCompletionEnd("student = Split()", "Split", 8, 13), 15);
            // 落点列 = `(` 的 1-based 显示列。**不能减 1** —— 减了就落进词中间
            // （`Spli|t`）。
            Eq("caret-col/括号前", Rules.DispCol("student = Split()", 15), 16);

            // SetCaret 的列语义（"第 N 个字符之前"）就靠这一个函数，
            //   少了这个 -1，真机上每个 SetCaret 都会稳定右偏一格。
            Eq("caret-col/第1个字符之前", Rules.CaretColForCharCol("    n = 1", 1), 1);
            Eq("caret-col/第5个字符之前", Rules.CaretColForCharCol("    n = 1", 5), 5);
            Eq("caret-col/行尾之后", Rules.CaretColForCharCol("    Dim abcd", 13), 13);
            Eq("caret-col/全角行", Rules.CaretColForCharCol("    学生 = 1", 7), 9);

            // VBE 给的是显示列。纯 ASCII 行两种语义数值相同（所以这个坑
            //   只在有中文/全角的行上暴露）；全角行必须换算，否则取词会落空。
            Eq("char-col/纯ascii 恒等", Rules.CharColFromDisp("student = Split()", 16), 16);
            Eq("char-col/全角缩进", Rules.CharColFromDisp("    学生 = Split()", 17), 15);
            Eq("char-col/行尾", Rules.CharColFromDisp("abc", 4), 4);
            Eq("char-col/超出钳到行尾", Rules.CharColFromDisp("abc", 99), 4);
            // 换算之后再取词 —— 这一对是"含中文的行到底弹不弹"的判据
            Eq("char-col/全角行取词", Rules.WordBefore("    学生 = Split()",
                Rules.CharColFromDisp("    学生 = Split()", 17), out _), "Split");

            FuzzyHit mh;
            Ok("match/前缀", Fuzzy.Match("MsgBox", "msg", out mh) && mh.First == 0 && mh.Kind == 3);
            Ok("match/边界跳步", Fuzzy.Match("selBegin", "sb", out mh) &&
                                 mh.First == 0 && mh.Span == 4 && mh.Kind == 2);
            Ok("match/非前缀单字不算前缀", Fuzzy.Match("abcDef", "b", out mh) &&
                                           mh.First == 1 && mh.Kind < 3);
            Ok("match/不命中", !Fuzzy.Match("abc", "xyz", out mh));

            Ok("resolved/精确且无更长",
                Rules.PrefixIsResolved("sub", new List<string> { "sub", "function" }));
            Ok("resolved/有更长候选不算打全",
                !Rules.PrefixIsResolved("sub", new List<string> { "sub", "subset" }));
            Ok("resolved/没精确命中",
                !Rules.PrefixIsResolved("su", new List<string> { "sub" }));
            // 单字符变量名（`Dim w As Worksheet`）：敲一个 `w` 就是"打全了"，
            //   照打全即收处理的话永远看不到候选 —— 所以一个字不判打全。
            Ok("resolved/单字符不算打全",
                !Rules.PrefixIsResolved("w", new List<string> { "w" }));
            Ok("resolved/单字符即便只有自己是候选也不收",
                !Rules.PrefixIsResolved("i", new List<string> { "i" }));
            Ok("resolved/两个字符打全仍然收",
                Rules.PrefixIsResolved("ab", new List<string> { "ab" }));

            // 成员位置判据（点号后整个让给 VBE）—— 认 / 不认的位置各钉几条。
            Ok("member-position/点号后", Rules.InMemberPosition("student.", 9));
            // 改口径：「词中间」也是成员位置（VBE 这时正在按已打的字过滤成员），
            //   旧判据只认"紧跟点号"，于是 `student.na|` 会弹出我们的裸名池。
            Ok("member-position/成员名打到一半", Rules.InMemberPosition("student.na", 10));
            Ok("member-position/空格后点号", Rules.InMemberPosition("a = b. ", 8));
            Ok("member-position/点号带空格", Rules.InMemberPosition("a.b .c", 6));
            Ok("member-position/) 之后", Rules.InMemberPosition("    Range(\"A1\").", 19));
            Ok("member-position/] 之后", Rules.InMemberPosition("    arr(0).Cou", 15));
            // With 块：点顶在行首 / 紧跟语句边界 —— VBE 同样会弹成员列表
            Ok("member-position/With 块行首点", Rules.InMemberPosition("    .Size", 10));
            Ok("member-position/With 块赋值点", Rules.InMemberPosition("    Set f = .Font", 17));
            Ok("member-position/语句分隔后的点", Rules.InMemberPosition("    a = 1: .Size", 15));
            // 反例：这些点号不是成员访问，必须照常提示
            Ok("member-position/小数点不是", !Rules.InMemberPosition("x = 1.5", 8));
            Ok("member-position/没有点号不是", !Rules.InMemberPosition("student", 8));
            Ok("member-position/运算符后不是", !Rules.InMemberPosition("x = a + b", 10));
            Ok("member-position/串字面量后不是", !Rules.InMemberPosition("x = \"abc\".c", 13));
            Ok("member-position/注释里的点不是", !Rules.InMemberPosition("x = 1 ' c.d", 12));
            Ok("member-position/字符串里的点不是", !Rules.InMemberPosition("s = \"a.b\"", 10));
            Ok("member-position/返回点号下标", Rules.MemberDotCol("student.na", 10) == 7);

            Ok("decl-position/Dim 名字", Rules.InDeclarationPosition("Dim cus", 8));
            // `As` 后面一律不弹（用户明确要求）—— 那里只可能是类型名
            Ok("decl-position/As 后不弹", Rules.InDeclarationPosition("Dim x As Inte", 13));
            Ok("decl-position/As 后空词也不弹", Rules.InDeclarationPosition("Dim x As ", 11));
            Ok("decl-position/Sub 名字", Rules.InDeclarationPosition("Sub Do", 7));
            Ok("decl-position/普通语句不是", !Rules.InDeclarationPosition("x = cus", 8));
            Ok("decl-position/Const 值位置", !Rules.InDeclarationPosition("Const PI = 3", 13));
            // 不带 Dim 的裸声明 —— 旧正则漏了这几条，于是 `Public glo|`
            //   被当成普通表达式，弹窗把"正在定义的那个名字自己"提示出来。
            Ok("decl-position/裸 Public", Rules.InDeclarationPosition("Public gCount", 13));
            Ok("decl-position/裸 Private", Rules.InDeclarationPosition("Private m_arr", 13));
            Ok("decl-position/裸 Global", Rules.InDeclarationPosition("Global g", 8));
            Ok("decl-position/裸 Public As", Rules.InDeclarationPosition("Public g As Long", 16));
            Ok("decl-position/Public Sub 不算裸声明误判", Rules.InDeclarationPosition("Public Sub DoIt(", 16));

            // `As` / `New` 之后一律不弹 —— 与 decl-position 互补：
            //   那条认"声明语句的行首"，这条认 `Set c = New |` / `For i As |` /
            //   过程参数 `a As |` 这些行首不是声明关键字的位置。
            Ok("type-kw/New 后（空词）", Rules.AfterTypeKeyword("    Set c = New ", 17));
            Ok("type-kw/New 后半截词", Rules.AfterTypeKeyword("    Set c = New cl", 17));
            Ok("type-kw/As 后（For 行）", Rules.AfterTypeKeyword("    For i As Lo", 14));
            Ok("type-kw/过程参数 As 后", Rules.AfterTypeKeyword("Sub f(a As Lo", 12));
            Ok("type-kw/New 大小写", Rules.AfterTypeKeyword("x = NEW cl", 9));
            // 反例：这几个都**不是**关键字位置，必须照常提示
            Ok("type-kw/空行不是", !Rules.AfterTypeKeyword("", 1));
            Ok("type-kw/普通语句不是", !Rules.AfterTypeKeyword("x = newValue", 5));
            Ok("type-kw/Renew 不算 New", !Rules.AfterTypeKeyword("x = Renew", 5));
            Ok("type-kw/As 出现在词中间不算", !Rules.AfterTypeKeyword("x = ClassName", 5));

            // 把"正在输入的词"挖掉（自我提示的根治办法）
            Eq("blank-token/挖掉当前词",
                Rules.BlankToken("Public glo\r\nx = 1", 1, 8, "glo"), "Public    \nx = 1");
            Eq("blank-token/位置对不上就不动",
                Rules.BlankToken("Public abc\r\nx = 1", 1, 8, "glo"), "Public abc\r\nx = 1");
            Eq("blank-token/中文词",
                Rules.BlankToken("    学生 = 1", 1, 5, "学生"), "      " + " = 1");

            Ok("comment/注释里", Rules.InCommentOrString("' abc", 5));
            Ok("comment/字符串里", Rules.InCommentOrString("x = \"ab\"", 8));
            Ok("comment/普通代码不是", !Rules.InCommentOrString("x = 1", 5));

            char ch;
            Ok("line-edit/插入一个字符", Rules.LineEditInserted("Dim ab", "Dim abc", out ch) && ch == 'c');
            Ok("line-edit/删除不算", !Rules.LineEditInserted("Dim abc", "Dim ab", out ch));
            Ok("line-edit/替换不算", !Rules.LineEditInserted("Dim ab", "Dim xb", out ch));

            Ok("call-site/赋值右侧", Rules.IsCallSite("x = getVa", 9));
            Ok("call-site/点号后不是", !Rules.IsCallSite("a.b", 3));
            Ok("call-site/声明位置不是", !Rules.IsCallSite("Sub foo", 5));
            Ok("call-site/As 之后不是", !Rules.IsCallSite("Dim x As Long", 10));
            // 行首（只有缩进）也是调用点 —— "函数提示词不补括号"就栽在这条上
            Ok("call-site/行首是", Rules.IsCallSite("    MsgB", 5));
            Ok("call-site/顶格行首是", Rules.IsCallSite("MsgB", 1));
            Ok("call-site/注释之后不是", !Rules.IsCallSite("Call MsgB", 6));
        }

        // ------------------------------------------------------------------
        private const string Sample = @"Option Explicit
Private mCount As Long

Public Sub Main()
    Dim a As Long, b As String
    a = 1
    Call Helper(a)
    Set obj = New Collection
End Sub

Private Function GetName(ByVal idx As Long, Optional ByVal tag As String = """") As String
    GetName = ""x""
End Function

Public Property Get Size() As Long
    Size = 1
End Property

Public Sub Main()
    ' 同名过程的第二个
End Sub

Public Enum Color2
    clRed = 1
    clBlue
End Enum

Public Type Point2
    X As Long
    Y As Long
End Type
";

        private static void TestParser()
        {
            var procs = Parser.ExtractProcedures(Sample);
            Eq("procs/数量", procs.Count, 4);
            Eq("procs/第一个名字", procs[0].Name, "Main");
            Eq("procs/第一个种类", procs[0].Kind, "sub");
            Eq("procs/函数种类", procs[1].Kind, "function");
            Eq("procs/属性种类", procs[2].Kind, "property");
            Ok("procs/同名第二个带序号", procs[3].ScopeKey == "Main#2", procs[3].ScopeKey);

            var p = Parser.ProcOfLine(procs, 6);
            Ok("procs/第 6 行在 Main 里", p != null && p.Name == "Main", p?.ToString());

            // ---- 单行过程 / 缺 End Sub 的过程边界 ----
            var sp1 = Parser.ExtractProcedures("Sub A(): x = 1: End Sub\r\nDim q As Long\r\n");
            Ok("procs/单行过程只占一行",
                sp1.Count == 1 && sp1[0].StartLine == 1 && sp1[0].EndLine == 1,
                sp1.Count + ":" + (sp1.Count > 0 ? sp1[0].StartLine + "-" + sp1[0].EndLine : ""));
            Ok("procs/单行过程之后的行不在过程里",
                Parser.ProcOfLine(sp1, 2) == null);

            var sp2 = Parser.ExtractProcedures("Sub A()\r\n    x = 1\r\nSub B()\r\n    y = 2\r\nEnd Sub\r\n");
            Ok("procs/缺 End Sub 时在下一个过程头前收尾",
                sp2.Count == 2 && sp2[0].EndLine == 2 && sp2[1].StartLine == 3 && sp2[1].EndLine == 5,
                string.Join(",", sp2.Select(x => x.Name + ":" + x.StartLine + "-" + x.EndLine)));

            Ok("option-explicit/有", Parser.OptionExplicit(Sample));
            Ok("option-explicit/没有", !Parser.OptionExplicit("Sub a()\r\nEnd Sub"));

            Eq("code-only/去注释", Parser.CodeOnly("x = 1 ' abc").Trim(), "x = 1");
            Eq("code-only/去字符串", Parser.CodeOnly("x = \"abc\"").Trim(), "x =");
            // Rem 之后填等长空格（保持长度守恒），所以断言用 Trim
            Eq("code-only/Rem", Parser.CodeOnly("Rem abc").Trim(), "");
            // `Rem` 只要在语句起点就是注释，不限于行首 ——
            //   `x = 1: Rem 说明` 若漏判，会把"说明"当代码收进候选池
            Eq("code-only/冒号后的 Rem", Parser.CodeOnly("x = 1: Rem 说明").Trim(), "x = 1:");
            Eq("code-only/不误伤 Remark", Parser.CodeOnly("Remark = 1").Trim(), "Remark = 1");
            Eq("code-only/不误伤 Rem2", Parser.CodeOnly("Rem2 = 1").Trim(), "Rem2 = 1");
        }

        // ------------------------------------------------------------------
        private static void TestEngine()
        {
            var eng = new Engine
            {
                Recent = new Recent(),
            };
            eng.PrimePool(Sample, "Module1");

            bool wp, ws;
            eng.Followup("Main", "Other", out wp, out ws);
            Ok("followup/sub 补空格", !wp && ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "Other", out wp, out ws);
            Ok("followup/function 补括号", wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "GetName", out wp, out ws);
            Ok("followup/函数体内给自己赋值不补", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "getname#2", out wp, out ws);
            Ok("followup/作用域键带序号也认得出自己", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("Size", "Other", out wp, out ws);
            Ok("followup/property 都不补", !wp && !ws, $"parens={wp} space={ws}");

            // 内建函数 / 关键字不再进候选，于是"按 Tab 之后补什么"
            //   只认自己写的过程；变量 / 类型名 / 内建词一律什么都不补。
            eng.Followup("mCount", "Other", out wp, out ws);
            Ok("followup/变量什么都不补", !wp && !ws, $"parens={wp} space={ws}");
            eng.Followup("Color2", "Other", out wp, out ws);
            Ok("followup/自定义类型名什么都不补", !wp && !ws, $"parens={wp} space={ws}");
            eng.Followup("MsgBox", "Other", out wp, out ws);
            Ok("followup/内建函数不再补括号", !wp && !ws, $"parens={wp} space={ws}");
            eng.Followup("Public", "Other", out wp, out ws);
            Ok("followup/关键字不再补空格", !wp && !ws, $"parens={wp} space={ws}");

            Eq("classify/工程过程", eng.Classify("Main"), "func");
            Eq("classify/工程属性", eng.Classify("Size"), "prop");
            Eq("classify/变量", eng.Classify("mCount"), "var");
            Eq("classify/不认识的名字按变量", eng.Classify("某个名字"), "var");

            // 排序：作用域更近的优先（同名同命中质量时）
            var a = new Candidate { Name = "Main", ScopeKind = 4, Pos = 0, Span = 2, Kind = "func" };
            var b = new Candidate { Name = "Main2", ScopeKind = 2, Pos = 0, Span = 2, Kind = "func" };
            Ok("sort/作用域近的优先", Engine.Cmp(a, b) < 0);

            var c1 = new Candidate { Name = "abc", ScopeKind = 3, Pos = 0, Span = 1 };
            var c2 = new Candidate { Name = "abc", ScopeKind = 3, Pos = 1, Span = 1 };
            Ok("sort/命中位置靠前优先", Engine.Cmp(c1, c2) < 0);

            var r1 = new Candidate { Name = "a", ScopeKind = 3, Pos = 0, Span = 1, RecentRank = 5 };
            var r2 = new Candidate { Name = "b", ScopeKind = 3, Pos = 0, Span = 1, RecentRank = 1 };
            Ok("sort/最近使用优先", Engine.Cmp(r1, r2) < 0);

            // 命中质量必须排在作用域**前面**：打 `msgbox` 时，一个工程内的
            //   长名字只是"松散命中"，绝不能因为它"工程内"就顶掉精确命中的 MsgBox。
            var loose = new Candidate { Name = "MsgBoxLocalHelper", ScopeKind = 4, HitKind = 0, Pos = 3, Span = 9 };
            var exact = new Candidate { Name = "MsgBox", ScopeKind = 2, HitKind = 3, Pos = 0, Span = 6 };
            Ok("sort/命中质量优先于作用域", Engine.Cmp(exact, loose) < 0);

            // 名字短的在前（长度降序会把长的排前，正好相反）
            var shortN = new Candidate { Name = "ab", ScopeKind = 3, HitKind = 3, Pos = 0, Span = 2 };
            var longN = new Candidate { Name = "abcdef", ScopeKind = 3, HitKind = 3, Pos = 0, Span = 8 };
            Ok("sort/名字短的在前", Engine.Cmp(shortN, longN) < 0);

            // ---- 启动脚本契约----
            // "双击窗口一闪、项目没启动"出过两次，两次都在 .bat 上：
            //   echo 文本里一个裸右括号就让 cmd 丢弃整个 if 块并退出 255；
            //   LF-only 换行让 goto 标签和多行 if 块失效；
            //   还有更根本的一条：start 之后既不验证也不提示，程序真没起来用户也看不到。
            // 这些是【文件形态】问题，只能查字节，别用源码字符串护栏（注释会喂饱它）。
            var root = Settings.Home;
            var bats = Directory.GetFiles(root, "*.bat", SearchOption.AllDirectories)
                .Where(p => !p.Contains(@"\obj\") && !p.Contains(@"\bin\"))
                .ToArray();
            Ok("bat/找到启动脚本", bats.Length >= 4, bats.Length + " 个");

            var badEol = new List<string>();
            var badParen = new List<string>();
            foreach (var f in bats)
            {
                var bytes = File.ReadAllBytes(f);
                int lfOnly = 0;
                for (int i = 0; i < bytes.Length; i++)
                {
                    if (bytes[i] != (byte)'\n') continue;
                    if (i == 0 || bytes[i - 1] != (byte)'\r') lfOnly++;
                }
                if (lfOnly > 0) badEol.Add(Path.GetFileName(f) + "(LF=" + lfOnly + ")");

                foreach (var line in File.ReadAllLines(f))
                {
                    var t = line.Trim();
                    if (!t.StartsWith("echo", StringComparison.OrdinalIgnoreCase)) continue;
                    if (t.IndexOf(')') >= 0) badParen.Add(Path.GetFileName(f) + ": " + t);
                }
            }
            Ok("bat/必须是 CRLF 换行", badEol.Count == 0, string.Join(", ", badEol));
            Ok("bat/echo 不许有裸右括号", badParen.Count == 0, string.Join(" | ", badParen));

            var runCs = Path.Combine(root, "run-cs.bat");
            var runCsText = File.Exists(runCs) ? File.ReadAllText(runCs) : "";
            Ok("bat/启动后验证进程存活", runCsText.Contains("goto died") && runCsText.Contains("tasklist"),
                "缺存活校验");
            Ok("bat/失败时打印日志尾部",
                runCsText.Contains("crash.log") && runCsText.Contains("VbeComplete.log"),
                "没有把日志给用户看");
            Ok("bat/不用 find（Git 的同名程序会抢 /i）",
                runCsText.Contains("findstr") && !runCsText.Contains("| find /i"),
                "still uses find");
        }

        // ------------------------------------------------------------------
        // 编辑手感（块结构 / 缩进 / 自动配对 / 退格接管）
        // ------------------------------------------------------------------
        // 模糊匹配：这一层的每个断言都对应一种真实的边界情况，
        // 改算法前先看一眼 —— 里面有一条是"贪心会漏判"的经典反例。
        // ------------------------------------------------------------------
        // 候选池口径：池子里**只有自己写的名字**。
        //   · 不弹：VBA 关键字 / 内建函数 / 内建类型 / 模块名 / 表名 / 窗体名 /
        //           类名 / 自定义类型名 / Type 成员
        //   · 要弹：过程名、参数名、局部变量、模块级变量、枚举成员
        //   强制声明（Option Explicit）与非强制声明两种模块各跑一遍，
        //   证明"收窄候选"没有把作用域判定带坏。
        // ------------------------------------------------------------------
        private static void TestPoolContents()
        {
            const string code =
                "Private Type Point\r\n" +
                "    X As Long\r\n" +
                "End Type\r\n" +
                "Public Enum Color2\r\n" +
                "    clRed = 1\r\n" +
                "End Enum\r\n" +
                "Private mCount As Long\r\n" +
                "\r\n" +
                "Public Sub DoIt(ByVal studentName As String)\r\n" +
                "    Dim localOnly As Long\r\n" +
                "    localOnly = Len(studentName)\r\n" +
                "    If localOnly > 0 Then\r\n" +
                "        MsgBox \"ok\"\r\n" +
                "    End If\r\n" +
                "    Sheet1.Range(\"A1\").Value = 1\r\n" +
                "    localOnly = 2\r\n" +
                "End Sub\r\n";

            foreach (bool explicitOn in new[] { true, false })
            {
                string text = explicitOn ? "Option Explicit\r\n" + code : code;
                var be = new FakeBackend
                {
                    Text = text,
                    // Sheet1 是个文档模块：它的名字不许进候选（口径里的"表名"）
                    Extra =
                    {
                        new ModuleScope { Module = "Sheet1", ComponentType = 100 },
                    },
                    Snap = new Snapshot { Line = 16, Text = "    ", Module = "Module1", CaretCol = 5 },
                };
                var eng = new Engine { Recent = new Recent(), Backend = be };
                var items = eng.Trigger(true, be.Snap) ? eng.Items.ToList() : new List<Candidate>();
                string dump = string.Join(",", items.Select(c => c.Name));
                string tag = explicitOn ? "explicit" : "no-explicit";

                foreach (var want in new[] { "DoIt", "studentName", "localOnly", "mCount", "clRed" })
                    Ok("pool/" + tag + " 保留自己的名字 " + want,
                        items.Any(x => x.Name == want), dump);

                foreach (var bad in new[] { "Point", "Color2", "X", "MsgBox", "Len",
                                            "If", "Then", "Dim", "Option", "String",
                                            "Long", "Sheet1", "Range", "Value" })
                    Ok("pool/" + tag + " 不弹 " + bad,
                        !items.Any(x => string.Equals(x.Name, bad, StringComparison.OrdinalIgnoreCase)),
                        dump);
            }
        }

        // ------------------------------------------------------------------
        // 单字符变量名：打全了也要提示
        //   `Dim w As Worksheet` 里敲一个 `w`，本来就已经是"打全了"。按"打全即收"
        //   处理的话，光标一停在 w 后面列表就没了 —— 用户看到的是"i 能提示、
        //   w 不能"（i 只是沾了 itemsArr 这类更长候选的光）。
        //   单字符名字没有"再补一遍"的余地（补上去还是同一个字符），所以不判打全。
        // ------------------------------------------------------------------
        private static void TestSingleCharName()
        {
            Func<string, string, int, Engine> mk = (code, line, caret) =>
            {
                var be = new FakeBackend
                {
                    Text = code,
                    Snap = new Snapshot { Line = 4, Text = line, Module = "Mod1", CaretCol = caret },
                };
                return new Engine { Recent = new Recent(), Backend = be };
            };

            const string oneChar =
                "Option Explicit\r\nSub Test()\r\n    Dim w As Worksheet, t As Object\r\n    w\r\nEnd Sub\r\n";
            var e1 = mk(oneChar, "    w", 6);
            bool shown1 = e1.Trigger(false, e1.Backend.Snapshot());
            var one = shown1 ? e1.Items.ToList() : new List<Candidate>();
            Ok("single-char/敲 w 就提示 w",
                one.Any(x => x.Name == "w"),
                "shown=" + shown1 + " items=[" + string.Join(",", one.Select(x => x.Name))
                + "] hide=" + Settings.Flag("HIDE_ON_COMPLETE", true));

            const string twoChar =
                "Option Explicit\r\nSub Test()\r\n    Dim ab As Long\r\n    ab\r\nEnd Sub\r\n";
            var e2 = mk(twoChar, "    ab", 7);
            bool shown2 = e2.Trigger(false, e2.Backend.Snapshot());
            var two = shown2 ? e2.Items.ToList() : new List<Candidate>();
            Ok("single-char/两个字符打全了照旧收",
                two.Count == 0,
                "shown=" + shown2 + " items=[" + string.Join(",", two.Select(x => x.Name)) + "]");
        }

        // ------------------------------------------------------------------
        private static void TestFuzzy()
        {
            FuzzyHit h;

            // 1) 整块连续优先：arr 命中 dataArr 的 Arr，而不是散着的 a..r..r
            Ok("fuzzy/连续块", Fuzzy.Match("dataArr", "arr", out h));
            Eq("fuzzy/连续块-kind", h.Kind, 1);
            Eq("fuzzy/连续块-位置", string.Join(",", h.Positions), "4,5,6");

            // 2) 词首缩略：ds -> dataSheet 的 D 与 S（中间没有连续的 ds）
            Ok("fuzzy/词首缩略", Fuzzy.Match("dataSheet", "ds", out h));
            Eq("fuzzy/词首缩略-kind", h.Kind, 2);
            Eq("fuzzy/词首缩略-位置", string.Join(",", h.Positions), "0,4");

            // 3) ts -> dataSheet：t 在词中、S 在词边界
            Ok("fuzzy/边界跳步", Fuzzy.Match("dataSheet", "ts", out h));
            Eq("fuzzy/边界跳步-位置", string.Join(",", h.Positions), "2,4");

            // 4) 前缀 / 完全相等
            Ok("fuzzy/前缀", Fuzzy.Match("dataSheet", "dat", out h));
            Eq("fuzzy/前缀-kind", h.Kind, 3);
            Ok("fuzzy/完全相等", Fuzzy.Match("dataSheet", "datasheet", out h));
            Eq("fuzzy/完全相等-kind", h.Kind, 4);
            Ok("fuzzy/大小写无关", Fuzzy.Match("dataSheet", "DATAS", out h));

            // 5) 贪心漏判的经典反例：
            //    第 3 个字符 t 若贪心跳到词边界 T(10)，后面就没有 s 了；
            //    正确解是 [0,1,2,3,4,15]。反向可行性上界就是为这条加的。
            Ok("fuzzy/子序列不许漏判", Fuzzy.Match("getSettingTitle3", "getse3", out h));
            Eq("fuzzy/子序列-位置", string.Join(",", h.Positions), "0,1,2,3,4,15");
            Eq("fuzzy/子序列-最长连续段", h.Run, 5);

            // 6) 跳步但"真的打过一段"：xlworkfaul -> xlWorkbookDefault
            Ok("fuzzy/长名字跳步", Fuzzy.Match("xlWorkbookDefault", "xlworkfaul", out h));
            Eq("fuzzy/长名字-最长连续段", h.Run, 6);

            // 7) 不匹配 / 长度不够 / 空 query
            Ok("fuzzy/不匹配", !Fuzzy.Match("dataSheet", "zzz", out h));
            Ok("fuzzy/比名字还长", !Fuzzy.Match("abc", "abcd", out h));
            Ok("fuzzy/空 query 全都要", Fuzzy.Match("abc", "", out h));
            Eq("fuzzy/空 query-位置个数", h.Positions.Length, 0);

            // 8) 中文名字：没有大小写/驼峰，整块连续照样命中
            Ok("fuzzy/中文前缀", Fuzzy.Match("学生列表", "学生", out h));
            Eq("fuzzy/中文-kind", h.Kind, 3);

            // 9) 下划线也是词边界（VBA 里 m_Name 很常见）
            Ok("fuzzy/下划线边界", Fuzzy.Match("m_studentName", "sn", out h));
            Eq("fuzzy/下划线边界-位置", string.Join(",", h.Positions), "2,9");

            // 10) 高亮切段：相邻同类必须合并成一段（否则一个名字要画十几个文本项）
            var segs = PopupForm.SplitByHits("dataSheet", new[] { 0, 4 });
            Eq("hits/段数", segs.Count, 4);
            Eq("hits/首段文字", segs[0].Key, "d");
            Ok("hits/首段命中", segs[0].Value);
            Eq("hits/次段文字", segs[1].Key, "ata");
            Ok("hits/次段未命中", !segs[1].Value);
            Eq("hits/M", segs[2].Key, "S");
            Eq("hits/尾段", segs[3].Key, "heet");
            Eq("hits/空名字", PopupForm.SplitByHits("", new[] { 0 }).Count, 0);
        }

        // ------------------------------------------------------------------
        // 作用域：哪些名字在当前位置**看得见**。这里每一条都是"提示里混进
        // 用不了的名字"的真 bug 的反面。
        // ------------------------------------------------------------------
        private const string ScopedSample = @"Option Explicit
Private mModuleLevel As Long
Public gPublic As Long

Private Sub A()
    Dim tmpA As Long
    Dim shared As String
    tmpA = 1
End Sub

Public Sub B(ByVal p1 As Long)
    Dim tmpB As Long
    shared = ""x""
End Sub

Public Sub C()
End Sub
";

        /// <summary>
        /// 过程外不带 Dim 的裸声明逗号多名 / Friend / 只有名字。
        /// 三种写法都要能解析出来。
        /// </summary>
        private const string MultiBareSample = @"Option Explicit
Public a, b As Long
Private c, d As String
Friend e As Long
Global f
";

        private static void TestScope()
        {
            var ms = Scope.Analyze("Mod1", ScopedSample);

            var va = ms.Visible("A");
            Ok("scope/A 看得见自己的局部", va.ContainsKey("tmpA"));
            Eq("scope/局部档位=4", va["tmpA"].Value, 4);
            Ok("scope/A 看得见模块级 Private", va.ContainsKey("mModuleLevel"));
            Eq("scope/模块级档位=3", va["mModuleLevel"].Value, 3);
            Ok("scope/A 看不见 B 的局部", !va.ContainsKey("tmpB"));
            Ok("scope/A 看得见本模块过程名", va.ContainsKey("B"));

            var vb = ms.Visible("B");
            Ok("scope/参数是局部", vb.ContainsKey("p1"));
            Eq("scope/参数档位=4", vb["p1"].Value, 4);
            Ok("scope/B 看不见 A 的局部", !vb.ContainsKey("tmpA"));
            // 这条是核心：B 里**引用**了 shared，但 shared 声明在 A 里
            //   —— 不能因为"B 里出现过这个词"就把 A 的局部漏给 B。
            Ok("scope/别处引用过的局部不漏过来", !vb.ContainsKey("shared"));

            var ex = ms.Exported();
            Ok("scope/Public 过程可导出", ex.ContainsKey("B") && ex.ContainsKey("C"));
            Ok("scope/Private 过程不导出", !ex.ContainsKey("A"));
            Ok("scope/Public 变量可导出", ex.ContainsKey("gPublic"));
            Ok("scope/Private 变量不导出", !ex.ContainsKey("mModuleLevel"));
            Ok("scope/局部一律不导出", !ex.ContainsKey("tmpA") && !ex.ContainsKey("p1"));

            // 引擎层面：池子里有 shared，但在 B 的位置不该出现在候选里
            var eng = new Engine { Recent = new Recent() };
            eng.PrimePool(ScopedSample, "Mod1");
            var items = enginesTrigger(eng, "    sha", 8);
            Ok("scope/引擎在 B 里不提示 A 的局部",
                items.All(x => x.Name.ToLowerInvariant() != "shared"),
                string.Join(",", items.Select(x => x.Name)));

            // 同一个词在它自己的过程里必须提示得到
            var items2 = enginesTrigger(eng, "    sha", 8, "A");
            Ok("scope/引擎在 A 里提示得到自己的局部",
                items2.Any(x => x.Name.ToLowerInvariant() == "shared"),
                string.Join(",", items2.Select(x => x.Name)));

            // ---- 过程外裸声明要支持逗号多名 ----
            //   旧正则要求首名后面紧跟 As/(/行尾，于是 `Public a, b As Long`
            //   整行匹配不上，a 和 b 双双进不了候选池。
            var mb = Scope.Analyze("Mod2", MultiBareSample);
            Func<string, Decl> pick = n => mb.Decls.FirstOrDefault(x => x.Name == n);
            Ok("scope/裸声明多名：a 收进来了", pick("a") != null);
            Ok("scope/裸声明多名：b 收进来了", pick("b") != null);
            Eq("scope/裸声明多名：b 拿到 As 后面的类型", pick("b")?.TypeName, "Long");
            // VBA 口径：`Public a, b As Long` 里 a 是 Variant，只有 b 是 Long
            Eq<string>("scope/裸声明多名：a 仍是 Variant", pick("a")?.TypeName, null);
            Ok("scope/裸声明多名：Private 覆盖到每个名字",
                pick("c") != null && pick("d") != null && pick("c").IsPrivate && pick("d").IsPrivate);
            Ok("scope/裸声明 Friend 不算 Private", pick("e") != null && !pick("e").IsPrivate);
            Ok("scope/裸声明只写名字也收", pick("f") != null);
            Ok("scope/裸声明进了可见集", mb.Visible(null).ContainsKey("a") && mb.Visible(null).ContainsKey("b"));
        }

        /// <summary>造一个假快照跑 Trigger（不碰 COM），返回候选名。</summary>
        private static List<Candidate> enginesTrigger(Engine eng, string line, int caretDisp,
            string proc = "B")
        {
            int ln = proc == "A" ? 7 : 13;          // 7 在 Sub A 里，13 在 Sub B 里
            var be = new FakeBackend
            {
                Text = ScopedSample,
                Snap = new Snapshot { Line = ln, Text = line, Module = "Mod1", CaretCol = caretDisp },
            };
            eng.Backend = be;
            if (!eng.Trigger(false, be.Snap)) return new List<Candidate>();
            return eng.Items.ToList();
        }

        /// <summary>只给引擎喂快照的最小后端：不碰 COM。</summary>
        private sealed class FakeBackend : IVbeSource
        {
            public string Text;
            public Snapshot Snap;
            public List<ModuleScope> Extra = new List<ModuleScope>();
            // 上一次 ApplyCompletion 收到的参数（自检要断言"替换了哪一段"）
            public int LastStartCol, LastEndCol;
            public string LastCompletion;
            public Snapshot Snapshot(bool verify = true) => Snap;
            public string ModuleText() => Text;
            public List<string> ComponentNames() => new List<string>();
            public List<ModuleScope> ProjectScopes(out int signature)
            {
                signature = 0;
                return Extra;
            }
            public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
                string completion, bool wantParens, bool wantSpace, string enclosingProc)
            {
                LastStartCol = wordStartCol;
                LastEndCol = endCol;
                LastCompletion = completion;
                return true;
            }
            public string EnclosingProcName() => null;
        }

        // ------------------------------------------------------------------
        // `.` 之后的成员：只在"类型写得明明白白"时才补，否则整个让给 VBE。
        // ------------------------------------------------------------------
        private const string MemberSample = @"Option Explicit
Private Sub Main()
    Dim c As Cls1
    Set c = New Cls1
    c.
End Sub

Public Sub DoIt()
End Sub
";

        private const string Cls1Sample = @"Public Sub Refresh()
End Sub

Private Sub Hidden()
End Sub

Public Name As String
";

        // ------------------------------------------------------------------
        // 中文标识符 + Option Explicit
        // ------------------------------------------------------------------
        private const string ChineseSample = @"Option Explicit
Private 模块计数 As Long

Public Sub 汇总(ByVal 起始行 As Long)
    Dim 合计 As Long
    合计 = 1
End Sub

Public Function 取总分() As Long
    取总分 = 0
End Function
";

        private const string NoExplicitSample = @"Private Sub 测试()
    学生 = 1
    总分 = 学生 + 2
End Sub
";

        private const string HostObjSample = @"Option Explicit
Private Sub 点击()
    CommandButton1.Caption = ""确定""
    Set 按钮 = Me.Controls.Add(""Forms.CommandButton.1"")
End Sub
";

        private static void TestChineseAndExplicit()
        {
            // ---- 中文过程名 / 变量名必须解析出来 ----
            var ms = Scope.Analyze("模块1", ChineseSample);
            Ok("cn/中文过程名是声明", ms.Decls.Any(d => d.Name == "汇总" && d.KindHint == "sub"),
                string.Join(",", ms.Decls.Select(d => d.Name)));
            Ok("cn/中文函数名", ms.Decls.Any(d => d.Name == "取总分" && d.KindHint == "function"));
            Ok("cn/中文模块变量", ms.Decls.Any(d => d.Name == "模块计数" && d.Proc == null));
            Ok("cn/中文参数是局部", ms.Decls.Any(d => d.Name == "起始行" && d.Proc != null));

            var vis = ms.Visible("汇总");
            Ok("cn/中文局部可见", vis.ContainsKey("合计"));
            Ok("cn/中文参数可见", vis.ContainsKey("起始行"));

            // 中文名要能进候选池并被模糊命中
            var cnSnap = new Snapshot
            {
                // CaretCol 是**显示列**：4 个空格 + 一个汉字 = 宽 6，光标落在 7
                Line = 6, Text = "    合", Module = "模块1", CaretCol = 7,
            };
            var cnBe = new FakeBackend { Text = ChineseSample, Snap = cnSnap };
            var eng = new Engine { Recent = new Recent(), Backend = cnBe };
            var items = eng.Trigger(false, cnSnap) ? eng.Items.ToList() : new List<Candidate>();
            Ok("cn/打中文首字能提示", items.Any(c => c.Name == "合计"),
                string.Join(",", items.Select(c => c.Name)));

            // ---- Option Explicit ----
            Ok("explicit/认出来了", Scope.Analyze("M", ChineseSample).OptionExplicit);
            Ok("explicit/没写就不算", !Scope.Analyze("M", NoExplicitSample).OptionExplicit);

            var noExp = Scope.Analyze("M", NoExplicitSample);
            Ok("explicit/无该选项时未声明名字进池",
                noExp.FreeIdents.Contains("学生") && noExp.FreeIdents.Contains("总分"),
                string.Join(",", noExp.FreeIdents));

            var expSample = "Option Explicit\r\nPrivate Sub T()\r\n    未定义 = 1\r\nEnd Sub\r\n";
            var exp = Scope.Analyze("M", expSample);
            Ok("explicit/有该选项时未声明名字不进池", !exp.FreeIdents.Contains("未定义"),
                string.Join(",", exp.FreeIdents));

            // ---- 自我提示：正在打的这个词本身不该进候选池 ----
            // 症状：定义全局/模块变量时（`Public glo|`），解析会把半截名当成
            //   已声明的名字，于是候选第一条就是"你正在打的这个词"，一弹一收像闪烁。
            const string selfSample = "Sub T()\r\n    myNewVar = 1\r\nEnd Sub\r\n";
            var selfBe = new FakeBackend
            {
                Text = selfSample,
                Snap = new Snapshot { Line = 2, Text = "    myNewVa", Module = "M", CaretCol = 13 },
            };
            var selfEng = new Engine { Recent = new Recent(), Backend = selfBe };
            bool selfShown = selfEng.Trigger(false, selfBe.Snap);
            Ok("self/正在输入的名字不提示自己",
                !selfShown || !selfEng.Items.Any(c => c.Name == "myNewVar"),
                selfShown ? string.Join(",", selfEng.Items.Select(c => c.Name)) : "(不弹)");

            // 宿主对象（点了点号的）在 Option Explicit 下仍要留
            var host = Scope.Analyze("M", HostObjSample);
            Ok("explicit/控件名（带点号）保留", host.FreeIdents.Contains("CommandButton1"),
                string.Join(",", host.FreeIdents));
            // `Set 按钮 = ...` 的"按钮"是**被赋值的变量**，不是宿主给的对象。
            //   Option Explicit 下没 Dim 它就是编译错误，不能放行（原来的"前面是 Set
            //   就算宿主对象"那条豁免已经删掉）。
            Ok("explicit/Set 的未声明变量不再放行（按钮）", !host.FreeIdents.Contains("按钮"),
                string.Join(",", host.FreeIdents));

            // ---- 成员名（点号后面的）绝不进自由标识符池 ----
            // 症状：输入 cell 会提示 cells（泄露）。
            //   他的代码里只有 `wsR.Cells(...)` —— 这是**成员访问**，不是自由变量。
            //   把点号后面的每个词都当标识符收进池子的话，全工程任何模块打
            //   `cell` 都会冒出 `Cells`。同理 `Value` / `Print` / `Range` 全在漏。
            const string memberSample =
                "Sub T()\r\n    n = wsR.Cells(r, 1).Value\r\nEnd Sub\r\n";
            var mem = Scope.Analyze("M", memberSample);
            Ok("成员名不进池（wsR.Cells / .Value 都不收）",
                !mem.FreeIdents.Contains("Cells") && !mem.FreeIdents.Contains("Value"),
                string.Join(",", mem.FreeIdents));
            Ok("对象名照收（wsR 后面跟点号 -> 是对象，要提示）",
                mem.FreeIdents.Contains("wsR"),
                string.Join(",", mem.FreeIdents));

            // 口径变了：`Cells(1,1)` 是**函数调用**，不再当"自动声明的变量"收。
            //   同一类现象 —— Worksheet 里打 `ra` 第一条弹 Excel 的 `Range`，
            //   把 `Rate` / `Randomize` 全顶到后面。宿主函数不进候选池。
            const string bareSample = "Sub T()\r\n    n = Cells(1, 1)\r\nEnd Sub\r\n";
            Ok("裸用 Cells(1,1) 不再按变量收调用不是自动声明）",
                !Scope.Analyze("M", bareSample).FreeIdents.Contains("Cells"),
                string.Join(",", Scope.Analyze("M", bareSample).FreeIdents));

            // ---- 引擎层面：Option Explicit 的模块里，未声明名字不该出现在候选 ----
            var eng2 = new Engine { Recent = new Recent() };
            var be = new FakeBackend
            {
                Text = expSample,
                Snap = new Snapshot { Line = 3, Text = "    未定", Module = "M", CaretCol = 7 },
            };
            eng2.Backend = be;
            bool shown = eng2.Trigger(false, be.Snap);
            Ok("explicit/候选里没有未声明名字",
                !shown || !eng2.Items.Any(c => c.Name == "未定义" || c.Name == "未定"),
                shown ? string.Join(",", eng2.Items.Select(c => c.Name)) : "(不弹)");
        }

        /// <summary>
        /// 加的两道护栏：
        ///   1. AcceptName 必须复核"现在还是算候选时那个位置吗"；
        ///   2. 内建不能被"同名局部变量"一票否决。
        /// </summary>
        private static void TestGuards()
        {
            // ---- 1. 候选算出来之后光标被挪走 ----
            // 候选窗开着时按 Home（或鼠标点到同一行的别处）：行文本没变，轮询不会
            // 重触发，列表里还是旧候选。拿**新**的取词位置去替换就会 ——
            // 按过 Home 就变成"把候选插到行首"，点到词中间就"把那个词截成两半"。
            const string sample = "Private Sub T()\r\n    Dim studentName As String\r\n    stud\r\nEnd Sub\r\n";
            var be = new FakeBackend
            {
                Text = sample,
                Snap = new Snapshot { Line = 3, Text = "    stud", Module = "M", CaretCol = 9 },
            };
            var eng = new Engine { Recent = new Recent(), Backend = be };
            Ok("guard/候选先弹出来", eng.Trigger(false, be.Snap),
                string.Join(",", eng.Items.Select(c => c.Name)));

            // 模拟"按 Home"：同一行，光标跑到行首
            be.Snap = new Snapshot { Line = 3, Text = "    stud", Module = "M", CaretCol = 1 };
            Ok("guard/光标挪到行首后 Accept 被拒", !eng.AcceptName("studentName"));

            // 正常路径：又打了一个字母再按 Tab，取词起点没变，必须照常接受
            be.Snap = new Snapshot { Line = 3, Text = "    stud", Module = "M", CaretCol = 9 };
            eng.Trigger(false, be.Snap);
            be.Snap = new Snapshot { Line = 3, Text = "    student", Module = "M", CaretCol = 12 };
            Ok("guard/继续打字后放行（取词起点没变）", eng.AcceptName("studentName"));

            // ---- 2. 别的过程的局部变量不许泄露到这里 ----
            // A 里 `Dim localOnly` 是 A 的局部；B 里打 loc 绝不能看到它
            // （看到就等于教用户写一行编译不过/运行不对的代码）。
            const string shadow = "Private Sub A()\r\n    Dim localOnly As String\r\nEnd Sub\r\n\r\n" +
                                  "Private Sub B()\r\n    loc\r\nEnd Sub\r\n";
            var be2 = new FakeBackend
            {
                Text = shadow,
                Snap = new Snapshot { Line = 6, Text = "    loc", Module = "M", CaretCol = 8 },
            };
            var eng2 = new Engine { Recent = new Recent(), Backend = be2 };
            var items2 = eng2.Trigger(false, be2.Snap) ? eng2.Items.ToList() : new List<Candidate>();
            Ok("guard/别的过程的局部变量不泄露", !items2.Any(c => c.Name == "localOnly"),
                string.Join(",", items2.Select(c => c.Name)));

            // ---- 3. 替换的是**整个标识符**（对标 IDEA）----
            // 光标停在 `MsgB|ox` 中间按 Tab：只换光标左边那一截会写成 `MsgBoxox`。
            const string mid = "Private Sub T()\r\n    MsgBox\r\nEnd Sub\r\n";
            var be3 = new FakeBackend
            {
                Text = mid,
                Snap = new Snapshot { Line = 2, Text = "    MsgBox", Module = "M", CaretCol = 8 },
            };
            var eng3 = new Engine { Recent = new Recent(), Backend = be3 };
            eng3.Trigger(false, be3.Snap);
            eng3.AcceptName("MsgBox");
            // "    MsgBox" 长 10：抓取区间应是 [5, 11) —— 整词，不是 [5, 8)
            Ok("guard/替换整个标识符（光标在词中间）",
                be3.LastStartCol == 5 && be3.LastEndCol == 11,
                "start=" + be3.LastStartCol + " end=" + be3.LastEndCol);

            // ---- 4. 注释里手动唤出也不弹（对标 IDEA）----
            var be4 = new FakeBackend
            {
                Text = "Private Sub T()\r\n    ' 说明文字\r\nEnd Sub\r\n",
                Snap = new Snapshot { Line = 2, Text = "    ' 说明文字", Module = "M", CaretCol = 11 },
            };
            var eng4 = new Engine { Recent = new Recent(), Backend = be4 };
            Ok("guard/注释里 Ctrl+Space 也不弹", !eng4.Trigger(true, be4.Snap));

            // ---- 5. 命中过多时截断到 240（对标 IDEA 的"列表有上限"）----
            var bulk = new StringBuilder("Private Sub T()\r\n");
            for (int i = 0; i < 300; i++)
                bulk.Append("    Dim name" + i.ToString("D3") + " As Long\r\n");
            bulk.Append("End Sub\r\n");
            var be5 = new FakeBackend
            {
                Text = bulk.ToString(),
                Snap = new Snapshot { Line = 302, Text = "    ", Module = "M", CaretCol = 5 },
            };
            var eng5 = new Engine { Recent = new Recent(), Backend = be5 };
            bool shown5 = eng5.Trigger(true, be5.Snap);
            Ok("guard/命中过多时列表截断到 240",
                shown5 && eng5.Items.Count <= 240 && eng5.TruncatedFrom > 240,
                "shown=" + shown5 + " items=" + eng5.Items.Count + " total=" + eng5.TruncatedFrom);
        }

        /// <summary>
        /// 无 Option Explicit 时，只被当**函数调用**过的名字（`Range("d2")`
        /// / `InputBox("x")`）不能算"自动声明的变量"。
        /// <para>典型现象：Worksheet 模块里打 `ra`，第一条是 Excel 的 `Range`
        /// —— 它其实只是自己代码里 `Range("d2")` 调用过一次的宿主函数，被全量
        /// 标识符收进池子后拿着工程名档位排到了 `Rate` / `Randomize` 前面。</para>
        /// </summary>
        private static void TestFreeIdents()
        {
            const string ws = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n" +
                              "    Set Target = Target.Cells(1, 1)\r\n" +
                              "    If Target.Address = Range(\"d2\").Address Then\r\n" +
                              "        setValidity ActiveCell, \"c\"\r\n" +
                              "    End If\r\n" +
                              "End Sub\r\n";
            var ms = Scope.Analyze("Sheet1", ws);
            Ok("free/调用型名字不算自动声明变量（Range 不进池）",
                !ms.FreeIdents.Contains("Range"), string.Join(",", ms.FreeIdents));
            Ok("free/无括号的调用仍按变量收（setValidity）",
                ms.FreeIdents.Contains("setValidity"), string.Join(",", ms.FreeIdents));
            Ok("free/已声明的参数不进 FreeIdents（Target）",
                !ms.FreeIdents.Contains("Target"), string.Join(",", ms.FreeIdents));

            // 同一个名字既被当函数调用、又在别处当值用过 ⇒ 仍然是变量
            // （数组下标 `arr(1) = 5` 与赋值 `arr = Array()` 同时出现）
            const string arrCase = "Sub T()\r\n    arr(1) = 5\r\n    arr = Array(1, 2)\r\nEnd Sub\r\n";
            var ms2 = Scope.Analyze("M", arrCase);
            Ok("free/既调用又赋值时保留", ms2.FreeIdents.Contains("arr"),
                string.Join(",", ms2.FreeIdents));

            // ---- 只写下标的数组变量（从不写 arr = ...）也要保留 ----
            const string arrOnly = "Sub T()\r\n    arr(1) = 5\r\n    n = arr(2)\r\nEnd Sub\r\n";
            Ok("free/只写下标的数组变量仍然保留",
                Scope.Analyze("M", arrOnly).FreeIdents.Contains("arr"),
                string.Join(",", Scope.Analyze("M", arrOnly).FreeIdents));

            // ---- 冒号后的 Rem 是注释，里面的词不是变量 ----
            const string remMid = "Sub T()\r\n    x = 1: Rem 说明文字\r\nEnd Sub\r\n";
            Ok("free/冒号后 Rem 里的词不进池",
                !Scope.Analyze("M", remMid).FreeIdents.Contains("说明文字"),
                string.Join(",", Scope.Analyze("M", remMid).FreeIdents));

            // ---- GoTo 标签 / 标签定义 / Implements 都不是变量 ----
            const string labels = "Sub T()\r\n    On Error GoTo 错误处理\r\n    GoTo 结束\r\n" +
                                  "错误处理:\r\n    Exit Sub\r\n结束:\r\nEnd Sub\r\n";
            var msL = Scope.Analyze("M", labels);
            Ok("free/GoTo 标签与标签定义都不进池",
                !msL.FreeIdents.Contains("错误处理") && !msL.FreeIdents.Contains("结束"),
                string.Join(",", msL.FreeIdents));
            const string impl = "Implements IFoo\r\n\r\nSub T()\r\nEnd Sub\r\n";
            Ok("free/Implements 的接口名不进池",
                !Scope.Analyze("M", impl).FreeIdents.Contains("IFoo"),
                string.Join(",", Scope.Analyze("M", impl).FreeIdents));

            // ---- Type 块里的成员不进裸名池 ----
            const string typeBlock = "Public Type Person\r\n    Name As String\r\n    Age As Long\r\nEnd Type\r\n" +
                                     "Sub T()\r\n    p.Age = 1\r\nEnd Sub\r\n";
            var msT = Scope.Analyze("M", typeBlock);
            Ok("free/Type 成员不进裸名池（Name / Age）",
                !msT.FreeIdents.Contains("Name") && !msT.FreeIdents.Contains("Age"),
                string.Join(",", msT.FreeIdents));

            // ---- 词库里有的名字不再产生 var 副本 ----
            const string builtins = "Sub T()\r\n    n1 = Left(s, 2)\r\n    n2 = Len(s)\r\nEnd Sub\r\n";
            var msB = Scope.Analyze("M", builtins);
            Ok("free/内置函数不重复进 FreeIdents（Left / Len）",
                !msB.FreeIdents.Contains("Left") && !msB.FreeIdents.Contains("Len"),
                string.Join(",", msB.FreeIdents));

            // ---- `New Cls` 的类名、数字字面量的碎片都不是变量 ----
            const string misc = "Sub T()\r\n    Set c1 = New Collection\r\n    h1 = &HFF\r\n    e1 = 1E10\r\nEnd Sub\r\n";
            var msN = Scope.Analyze("M", misc);
            Ok("free/New 类名与数字字面量碎片不进池",
                !msN.FreeIdents.Contains("Collection") && !msN.FreeIdents.Contains("HFF") &&
                !msN.FreeIdents.Contains("E10"),
                string.Join(",", msN.FreeIdents));

            // ---- `Option Base 1` 里的词不是变量；`Let arr(1) = 5` 的 arr 是 ----
            var msO = Scope.Analyze("M", "Option Base 1\r\n\r\nSub T()\r\n    Let arr(1) = 5\r\nEnd Sub\r\n");
            Ok("free/Option 行里的词不进池",
                !msO.FreeIdents.Contains("Base"), string.Join(",", msO.FreeIdents));
            Ok("free/Let 数组赋值仍算变量", msO.FreeIdents.Contains("arr"),
                string.Join(",", msO.FreeIdents));

            // ---- 宿主对象名 / 枚举常量不进池（对齐 PY 的两道闸）----
            const string hostSample = "Sub T()\r\n" +
                                      "    Set rng = Range(\"A1\")\r\n" +
                                      "    n1 = Cells(1, 1).Value\r\n" +
                                      "    If n1 = xlUp Then n1 = vbYes\r\n" +
                                      "    Set ws = ActiveSheet\r\n" +
                                      "End Sub\r\n";
            var msH = Scope.Analyze("M", hostSample);
            Ok("host/Range / Cells / ActiveSheet 不进池",
                !msH.FreeIdents.Contains("Range") && !msH.FreeIdents.Contains("Cells") &&
                !msH.FreeIdents.Contains("ActiveSheet"),
                string.Join(",", msH.FreeIdents));
            Ok("host/枚举常量 xlUp / vbYes 不进池",
                !msH.FreeIdents.Contains("xlUp") && !msH.FreeIdents.Contains("vbYes"),
                string.Join(",", msH.FreeIdents));
            Ok("host/赋值目标 rng / ws 照常收",
                msH.FreeIdents.Contains("rng") && msH.FreeIdents.Contains("ws"),
                string.Join(",", msH.FreeIdents));

            // 例外：在**本模块里赋过值**的同名照常按变量收（`range = 1` 这种怪用法）
            var msH2 = Scope.Analyze("M", "Sub T()\r\n    range = 1\r\n    n2 = range * 2\r\nEnd Sub\r\n");
            Ok("host/被赋过值的同名照常收", msH2.FreeIdents.Contains("range"),
                string.Join(",", msH2.FreeIdents));

            // ---- 命名参数（`f(Name:=1)`）不是变量 ----
            var msNA = Scope.Analyze("M", "Sub T()\r\n    MsgBox Prompt:=x1, Buttons:=2\r\nEnd Sub\r\n");
            Ok("free/命名参数不进池（Prompt / Buttons）",
                !msNA.FreeIdents.Contains("Prompt") && !msNA.FreeIdents.Contains("Buttons"),
                string.Join(",", msNA.FreeIdents));

            // ---- 声明过的 `Set x = ...` 目标靠 Decls 进池（不靠 Set 豁免） ----
            var msSet = Scope.Analyze("M",
                "Option Explicit\r\nSub T()\r\n    Dim b As Object\r\n    Set b = CreateObject(\"X\")\r\nEnd Sub\r\n");
            Ok("explicit/声明过的 Set 目标仍在 Decls",
                msSet.Decls.Any(d => d.Name == "b"), string.Join(",", msSet.FreeIdents));

            // ---- 同名隐式变量在**两个过程里各有一份**，各自都要可见 ----
            //   （getTitle / getTitle2 里都有 `Set dic = ...`，只记第一次出现的
            //     过程的话，getTitle2 里就提示不出 dic）
            const string twoDic = "Function A()\r\n    Set dic = CreateObject(\"x\")\r\nEnd Function\r\n\r\n" +
                                  "Function B()\r\n    Set dic = CreateObject(\"x\")\r\n    di\r\nEnd Function\r\n";
            var msD = Scope.Analyze("M", twoDic);
            Ok("free/同名隐式变量在两个过程都可见",
                msD.Visible("A").ContainsKey("dic") && msD.Visible("B").ContainsKey("dic"),
                "A=" + string.Join(",", msD.Visible("A").Keys) +
                " | B=" + string.Join(",", msD.Visible("B").Keys));

            // ---- 隐式变量只在"首次出现的过程"里可见（PY 的 scope="proc"）----
            const string twoProcs = "Sub A()\r\n    num = 1\r\nEnd Sub\r\n\r\nSub B()\r\n    nu\r\nEnd Sub\r\n";
            var msP = Scope.Analyze("M", twoProcs);
            Ok("free/隐式变量在本过程可见", msP.Visible("A").ContainsKey("num"),
                string.Join(",", msP.Visible("A").Keys));
            Ok("free/隐式变量不跨过程泄露（B 里看不见 A 的 num）",
                !msP.Visible("B").ContainsKey("num"),
                string.Join(",", msP.Visible("B").Keys));

            // ---- 隐式变量是**模块私有**的，不能跨模块泄露 ----
            var otherA = Scope.Analyze("A", "Sub T()\r\n    foo = 1\r\nEnd Sub\r\n");
            var beCross = new FakeBackend
            {
                Text = "Sub W()\r\n    fo\r\nEnd Sub\r\n",
                Snap = new Snapshot { Line = 2, Text = "    fo", Module = "B", CaretCol = 7 },
                Extra = { otherA },
            };
            var engCross = new Engine { Recent = new Recent(), Backend = beCross };
            var itemsCross = engCross.Trigger(false, beCross.Snap)
                ? engCross.Items.ToList() : new List<Candidate>();
            Ok("free/别的模块的隐式变量不泄露",
                !itemsCross.Any(c => c.Name == "foo"),
                string.Join(",", itemsCross.Select(c => c.Name)));

            // 同一个名字在**本模块**里写过，就必须照常提示（证明上面不是"什么都没弹"）
            var beSame = new FakeBackend
            {
                Text = "Sub W()\r\n    foo = 1\r\n    fo\r\nEnd Sub\r\n",
                Snap = new Snapshot { Line = 3, Text = "    fo", Module = "B", CaretCol = 7 },
            };
            var engSame = new Engine { Recent = new Recent(), Backend = beSame };
            var itemsSame = engSame.Trigger(false, beSame.Snap)
                ? engSame.Items.ToList() : new List<Candidate>();
            Ok("free/本模块的隐式变量照常提示",
                itemsSame.Any(c => c.Name == "foo"),
                string.Join(",", itemsSame.Select(c => c.Name)));
        }

        private static void TestMember()
        {
            var cls = Scope.Analyze("Cls1", Cls1Sample);
            Eq("member/类型解析", cls.TypeOf("c", null), null);      // 模块级没有 c
            var mod = Scope.Analyze("Mod1", MemberSample);
            Eq("member/Dim As 收类型", mod.TypeOf("c", "Main"), "Cls1");

            // 默认口径 —— 点号之后**一个候选都不出**，整个让给 VBE 的
            //   成员列表。这条必须先钉住：下面
            //   那些"收窄"断言都跑在 MEMBER_POPUP=1 下，不先钉默认口径的话，
            //   一旦判据写错（把 `c.` 判成非成员位置），整套测试都不会红。
            var def = new Engine { Recent = new Recent() };
            var beDef = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    c.", Module = "Mod1", CaretCol = 7 },
            };
            def.Backend = beDef;
            Ok("member/默认点号后不弹（即使类型已知）", !def.Trigger(false, beDef.Snap));
            var beDef2 = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    c.Ref", Module = "Mod1", CaretCol = 10 },
            };
            def.Backend = beDef2;
            Ok("member/默认点号后打了半截词也不弹", !def.Trigger(false, beDef2.Snap));

            // ---- 以下都是 MEMBER_POPUP=1 那条路（用户显式要候选）----
            var eng = new Engine { Recent = new Recent(), MemberPopup = true };
            var be = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    c.", Module = "Mod1", CaretCol = 7 },
            };
            eng.Backend = be;
            bool shown = eng.Trigger(false, be.Snap);
            Ok("member/能确定类型就出成员", shown, "trigger=false");
            var names = eng.Items.Select(x => x.Name).ToList();
            Ok("member/含 Public 过程", names.Contains("Refresh"), string.Join(",", names));
            Ok("member/含 Public 变量", names.Contains("Name"), string.Join(",", names));
            Ok("member/不含 Private 成员", !names.Contains("Hidden"), string.Join(",", names));
            Ok("member/不含全工程噪音", !names.Contains("Main"), string.Join(",", names));

            // 类型不确定（变量从没声明过）=> 退回全工程候选（口径：
            // 定不出类型时宁可给宽的，也不给"收窄成空"）
            eng.Hide();
            var be2 = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    obj.", Module = "Mod1", CaretCol = 9 },
            };
            eng.Backend = be2;
            bool shown2 = eng.Trigger(false, be2.Snap);
            Ok("member/类型不确定时退回全工程候选",
                shown2 && eng.Items.Any(x => x.Name == "DoIt"),
                shown2 ? string.Join(",", eng.Items.Take(8).Select(x => x.Name)) : "(不弹)");

            // 点号后面已经开始打词：收窄仍然生效
            eng.Hide();
            var be3 = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    c.Ref", Module = "Mod1", CaretCol = 10 },
            };
            eng.Backend = be3;
            Ok("member/点号后继续打字仍收窄", eng.Trigger(false, be3.Snap));
            Ok("member/收窄后只剩一个",
                eng.Items.Count == 1 && eng.Items[0].Name == "Refresh",
                string.Join(",", eng.Items.Select(x => x.Name)));

            // `Me.` = 当前模块自己的 Public 成员
            eng.Hide();
            var be4 = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    Me.", Module = "Mod1", CaretCol = 8 },
            };
            eng.Backend = be4;
            Ok("member/Me 用当前模块", eng.Trigger(false, be4.Snap));

            // ---- `Me.` 里 Private 成员也算 ----
            //   套 Exported() 的话，那是"别的模块看得见什么"，正好把
            //   Private 全滤掉了；而 `Me.PrivateSub` 完全合法。
            var meNames = eng.Items.Select(x => x.Name).ToList();
            Ok("member/Me 含自己的 Private 过程", meNames.Contains("Main"), string.Join(",", meNames));
            Ok("member/Me 含自己的 Public 过程", meNames.Contains("DoIt"), string.Join(",", meNames));
            Ok("member/Me 不含过程局部", !meNames.Contains("c"), string.Join(",", meNames));

            // ---- 注释里（' 和 Rem）一律不弹 ----
            //   旧闸门只认 `'` 和字符串，`Rem Do |` 里照样弹全工程候选，
            //   按 Tab 还能把这行注释改写掉。这里挑 `Do` 是因为它有真候选
            //   （DoIt），否则"不弹"可能只是"没匹配"的假象。
            eng.Hide();
            var beRem = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 7, Text = "Rem Do", Module = "Mod1", CaretCol = 7 },
            };
            eng.Backend = beRem;
            Ok("gate/Rem 注释里不弹候选", !eng.Trigger(false, beRem.Snap));

            eng.Hide();
            var beTick = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 7, Text = "  ' Do", Module = "Mod1", CaretCol = 7 },
            };
            eng.Backend = beTick;
            Ok("gate/撇号注释里不弹候选", !eng.Trigger(false, beTick.Snap));
            // 反向对照：同样的词在正常代码位置上是要弹的（否则上面两条没意义）
            eng.Hide();
            var beCode = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 7, Text = "    Do", Module = "Mod1", CaretCol = 7 },
            };
            eng.Backend = beCode;
            Ok("gate/同一行代码位置照常弹", eng.Trigger(false, beCode.Snap));
        }

        // ------------------------------------------------------------------
        private static void TestEditing()
        {
            // ---- 注释/解除注释之后的光标列----
            // 注释在行首插一个 `'`：光标前多一个字符 → 字符列 +1
            Eq("comment-caret/注释时右移一格", Editing.AfterCommentCharCol("    n = 1", 5, false), 6);
            Eq("comment-caret/注释时行首也右移", Editing.AfterCommentCharCol("    n = 1", 1, false), 2);
            // 解除时删掉的是行首那个 `'`：只有它在光标之前才左移
            Eq("comment-caret/解除时左移一格", Editing.AfterCommentCharCol("'    n = 1", 6, true), 5);
            Eq("comment-caret/撇号在光标之后不动",
                Editing.AfterCommentCharCol("    n = 'x'", 5, true), 5);

            // ---- 基础换算 ----
            Eq("indent/4 空格", Editing.IndentWidth("    x"), 4);
            Eq("indent/tab 算 4 列", Editing.IndentWidth("\tx"), 4);
            Eq("leading-ws/混合", Editing.LeadingWs(" \t x"), " \t ");

            // ---- 块结构识别 ----
            Eq("block/If Then", Editing.BlockKind("If x Then"), "if");
            Eq("block/End If 不是块头", Editing.BlockKind("End If"), null);
            Eq("block/For", Editing.BlockKind("    For i = 1 To 3"), "for");
            Eq("block/Do While", Editing.BlockKind("Do While x"), "do");
            Eq("block/While", Editing.BlockKind("While x"), "while");
            Eq("block/With", Editing.BlockKind("With rng"), "with");
            Eq("block/Select Case", Editing.BlockKind("Select Case x"), "select");
            Eq("block/Select 少了 Case", Editing.BlockKind("Select x"), null);
            Eq("block/Private Sub", Editing.BlockKind("Private Sub Foo()"), "sub");
            Eq("block/Static Function", Editing.BlockKind("Static Function F() As Long"), "function");
            Eq("block/Property Get", Editing.BlockKind("Property Get X()"), "property");
            Eq("block/Type", Editing.BlockKind("Type Foo"), "type");
            Eq("block/Enum", Editing.BlockKind("Enum Color"), "enum");
            Eq("block/Else 是分支", Editing.BlockKind("Else"), "branch");
            Eq("block/ElseIf 带 Then", Editing.BlockKind("ElseIf y Then"), "branch");
            Eq("block/ElseIf 没 Then", Editing.BlockKind("ElseIf y"), null);
            Eq("block/Case 是分支", Editing.BlockKind("Case 1"), "branch");
            Eq("block/条件编译", Editing.BlockKind("#If VBA7 Then"), "directive");
            Eq("block/单选 If", Editing.BlockKind("If x Then y = 1"), null);
            Eq("block/Next 不是块头", Editing.BlockKind("Next i"), null);
            Eq("block/普通赋值", Editing.BlockKind("x = 1"), null);
            Eq("block/续行符结尾", Editing.BlockKind("If x Then _"), null);
            Eq("block/字符串里的 Then 不算", Editing.BlockKind("If x = \"a Then\" Then"), "if");
            Eq("block/Rem 整行注释", Editing.BlockKind("Rem If x Then"), null);
            Eq("block/整行注释不是块头", Editing.BlockKind("' For i = 1"), null);

            Eq("closer/For -> Next", Editing.BlockCloser("For i = 1 To 3"), "Next");
            Eq("closer/Do -> Loop", Editing.BlockCloser("Do"), "Loop");
            Eq("closer/With -> End With", Editing.BlockCloser("With rng"), "End With");
            Eq("closer/Else 没有收尾", Editing.BlockCloser("Else"), null);
            Eq("closer/#If 没有收尾", Editing.BlockCloser("#If VBA7 Then"), null);

            Ok("closer-line/Next i", Editing.IsBlockCloserLine("Next i"));
            Ok("closer-line/End If 带注释", Editing.IsBlockCloserLine("End If  '结束"));
            Ok("closer-line/For 不是收尾", !Editing.IsBlockCloserLine("For i = 1 To 3"));

            // ---- 缩进推算 ----
            Eq("indent/块头深一级", Editing.NextLineIndent("    For i = 1 To 3", new List<string>(), "    "),
                "        ");
            Eq("indent/普通行抄自己", Editing.NextLineIndent("    x = 1", new List<string>(), "    "), "    ");
            Eq("indent/注释跟上方第一个非注释行",
                Editing.NextLineIndent("    ' 注释", new List<string> { "Sub A()" }, "    "), "    ");
            Eq("indent/找不到参考行就用自己",
                Editing.NextLineIndent("    ' 注释", new List<string>(), "    "), "    ");
            Eq("indent/Tab 工程跟着用 Tab",
                Editing.NextLineIndent("\tIf x Then", new List<string>(), "\t"), "\t\t");

            // ---- 分支对齐 ----
            string mine, next;
            Ok("branch/Else 与 If 齐平",
                Editing.BranchAlign("    Else", new List<string> { "    If x Then" }, "    ",
                    out mine, out next) && mine == "    " && next == "        ");
            Ok("branch/Case 比 Select Case 深一级",
                Editing.BranchAlign("    Case 1", new List<string> { "    Select Case x" }, "    ",
                    out mine, out next) && mine == "        " && next == "            ");
            Ok("branch/找不到归属者就不动",
                !Editing.BranchAlign("    Else", new List<string>(), "    ", out mine, out next));
            Ok("branch/#Else 与 #If 齐平",
                Editing.BranchAlign("#Else", new List<string> { "#If VBA7 Then" }, "    ",
                    out mine, out next) && mine == "" && next == "    ");

            // ---- 回车接管 ----
            string blockLine = "    For i = 1 To 3";
            Ok("enter/块头行尾接管",
                Editing.EnterIndentWanted(blockLine, Rules.DispCol(blockLine, blockLine.Length)));
            Ok("enter/普通行不接管",
                !Editing.EnterIndentWanted("    x = 1", Rules.DispCol("    x = 1", 9)));
            string comment = "    ' 注释";
            Ok("enter/整行注释接管",
                Editing.EnterIndentWanted(comment, Rules.DispCol(comment, comment.Length)));
            Ok("enter/光标不在行尾不接管", !Editing.EnterIndentWanted(blockLine, 10));
            string tabLine = "\tIf x Then";
            Ok("enter/含 Tab 不接管",
                !Editing.EnterIndentWanted(tabLine, Rules.DispCol(tabLine, tabLine.Length)));
            string wideLine = "    For 每个 = 1 To 10";
            Ok("enter/全角行只认行尾",
                Editing.EnterIndentWanted(wideLine, Rules.DispCol(wideLine, wideLine.Length)) &&
                !Editing.EnterIndentWanted(wideLine, Rules.DispCol(wideLine, wideLine.Length - 3)));

            // ---- 退格 ----
            Ok("bs/整行空白接管", Editing.BlankLineDeleteWanted("    ", 5));
            Ok("bs/整行空白中间不接管", !Editing.BlankLineDeleteWanted("    ", 3));
            Ok("bs/真空行不接管", !Editing.BlankLineDeleteWanted("", 1));
            Ok("bs/有内容的行不接管", !Editing.BlankLineDeleteWanted("    x", 6));

            int left, right;
            Ok("bs-pair/空括号中间接管",
                Editing.PairDeleteSpan("x = f()", Rules.DispCol("x = f()", 6), out left, out right) &&
                left == 5 && right == 6);
            Ok("bs-pair/不在括号中间不接管",
                !Editing.PairDeleteSpan("x = f()", Rules.DispCol("x = f()", 8), out left, out right));
            Ok("bs-pair/括号里有内容不接管",
                !Editing.PairDeleteSpan("f(a)", Rules.DispCol("f(a)", 2), out left, out right));
            Ok("bs-pair/空引号中间接管",
                Editing.PairDeleteSpan("s = \"\"", Rules.DispCol("s = \"\"", 5), out left, out right) &&
                left == 4 && right == 5);

            // ---- 自动配对 ----
            int caretOff;
            bool paired;
            string got = Editing.PairInsertion("abc", 3, '(', ')', out caretOff, out paired);
            Ok("pair/行尾补一对", got == "abc()" && caretOff == 4 && paired);
            got = Editing.PairInsertion("abc def", 0, '(', ')', out caretOff, out paired);
            Ok("pair/右侧还有内容只插左括号", got == "(abc def" && !paired);
            got = Editing.PairInsertion("foo(bar", 7, '(', ')', out caretOff, out paired);
            Ok("pair/实参表里仍补一对", got == "foo(bar()" && paired);
            got = Editing.PairInsertion("s = ", 4, '"', '"', out caretOff, out paired);
            Ok("pair/引号补一对", got == "s = \"\"" && paired);
            got = Editing.PairInsertion("s = \"ab", 7, '"', '"', out caretOff, out paired);
            Ok("pair/未闭合字符串里只补收尾", got == "s = \"ab\"" && !paired);
            got = Editing.PairInsertion("s = \"\"", 5, '"', '"', out caretOff, out paired);
            Ok("pair/右边已是引号就跨过去", got == null && caretOff == 6);
            // 注：`)` 的"跨过去"在 VbeBackend.InsertPair 里先挡掉了（那个分支要读
            //     COM 才知道右边是什么），PairInsertion 不负责它 —— 这里只验引号那一档。

            Eq("pair/全角括号收符号", Editing.CloseOf('（'), '）');
            Eq("pair/全角引号收符号", Editing.CloseOf('“'), '”');
            Eq("pair/不认识的符号", Editing.CloseOf('['), '\0');

            // ---- 配对的闸门：只挡"注释里"，不挡"字符串里" ----
            // 这一条是 `s = "abc|"` 里再按 `"` 别变成 `"abc""` 的前提：
            // 闸门要是把"在字符串里"也算进去，PairInsertion 的"跨过去 / 只补
            // 收尾"两档就永远走不到。
            Ok("gate/未闭合字符串不算注释", !Rules.InCommentOnly("s = \"abc", 9));
            Ok("gate/闭合字符串之后不算注释", !Rules.InCommentOnly("s = \"abc\"", 10));
            Ok("gate/同一行与 InCommentOrString 的差别",
                Rules.InCommentOrString("s = \"abc", 9) && !Rules.InCommentOnly("s = \"abc", 9));
            Ok("gate/引号后是注释", Rules.InCommentOnly("s = \"abc\" ' hi", 14));
            Ok("gate/行首撇号", Rules.InCommentOnly("' hi", 5));
            Ok("gate/字符串里的撇号不是注释", !Rules.InCommentOnly("MsgBox \"it's\"", 13));
            Ok("gate/Rem 整行注释", Rules.InCommentOnly("Rem 说明", 6));
            Ok("gate/Remember 不是注释", !Rules.InCommentOnly("Remember = 1", 11));
            Ok("gate/冒号后的 Rem", Rules.InCommentOnly("x = 1 : Rem 说明", 14));
            Ok("gate/语句中间的 Rem 不是注释", !Rules.InCommentOnly("x = Rem", 7));

            // ---- 弹窗死区`'` / `Rem` / 字符串三者都算 ----
            //   两条老判据各管一半，引擎原来只用其中一条，于是 Rem 里照弹。
            Ok("gate/死区：Rem 注释", Rules.InDeadZone("Rem 说明", 6));
            Ok("gate/死区：撇号注释", Rules.InDeadZone("x = 1 ' hi", 11));
            Ok("gate/死区：闭合字符串里", Rules.InDeadZone("s = \"ab\"", 8));
            Ok("gate/死区：未闭合字符串里", Rules.InDeadZone("s = \"abc", 9));
            Ok("gate/死区：普通代码不是", !Rules.InDeadZone("x = 1", 5));
            Ok("gate/死区：Remember 不是", !Rules.InDeadZone("Remember = 1", 11));

            Eq("call-paren/foo( 认得出", Editing.EnclosingCallParen("foo(", 4), 3);
            Eq("call-paren/运算分组 -> -1", Editing.EnclosingCallParen("x = (a + ", 9), -1);
            Eq("call-paren/字符串里的括号不算", Editing.EnclosingCallParen("s = \"(\" & foo(", 14), 13);

            Ok("string/未闭合", Editing.UnterminatedString("x = \"abc"));
            Ok("string/已闭合", !Editing.UnterminatedString("x = \"abc\""));
            Ok("string/注释里的引号不算", !Editing.UnterminatedString("x = 1  ' 他说\"你好"));
            Ok("string/双写引号是转义", !Editing.UnterminatedString("x = \"a\"\"b\""));

            // ---- 补收尾的完整口径 ----
            // 为什么不能只按"同级一律不补"处理：先写一个 For（回车补了 Next），
            //   再回到它上面补一个新的 For 按回车 —— 新 For 下面压着同级的代码，
            //   于是 Next 永远补不出来。
            var noAbove = new List<string>();
            Ok("closer-needed/下面是 End Sub 就补",
                Editing.CloserNeeded(new List<string> { "End Sub" }, "Next", 4, "for", noAbove));
            Ok("closer-needed/下面已挂着本块的 Next 不补",
                !Editing.CloserNeeded(new List<string> { "    Next" }, "Next", 4, "for", noAbove));
            Ok("closer-needed/块里已有内容且已有 End If 不补",
                !Editing.CloserNeeded(new List<string> { "        x = 1", "    End If" }, "End If", 4, "if", noAbove));
            Ok("closer-needed/同级已有代码压着也补（新 For 压在旧 For 上）",
                Editing.CloserNeeded(
                    new List<string> { "    For j = 1 To 2", "        x = 1", "    Next j" },
                    "Next", 4, "for", noAbove));
            Ok("closer-needed/同级普通语句也补（过程中间插一个块）",
                Editing.CloserNeeded(new List<string> { "    y = 2" }, "Next", 4, "for", noAbove));
            Ok("closer-needed/空行与注释跳过",
                Editing.CloserNeeded(new List<string> { "", "    ' 注释", "End Sub" }, "Next", 4, "for", noAbove));
            Ok("closer-needed/更浅的 Next 属于外层 For -> 补",
                Editing.CloserNeeded(new List<string> { "    Next" }, "Next", 8, "for",
                    new List<string> { "Sub Foo()", "    For i = 1 To 10" }));
            Ok("closer-needed/更浅的 Next 没人认领 -> 不补（防双份）",
                !Editing.CloserNeeded(new List<string> { "    Next" }, "Next", 8, "for",
                    new List<string> { "Sub Foo()" }));
            Ok("closer-needed/Else 不是块结束，继续往下找 End If",
                !Editing.CloserNeeded(
                    new List<string> { "    Else", "        y = 1", "    End If" },
                    "End If", 4, "if", noAbove));
            Ok("closer-needed/顶格行标签与 #If 指令整行跳过",
                Editing.CloserNeeded(new List<string> { "1:", "#If VBA7 Then", "    End Sub" },
                    "Next", 4, "for", noAbove));
            Ok("closer-needed/简化口径（不传 kind）保持不变",
                !Editing.CloserNeeded(new List<string> { "    y = 2" }, "Next", 4));

            // ---- 回车接管用的"块头前缀"宽松判据----
            // 快照落后一格时（`For i = 1 To ` 少个 3 / `If x > 0 The` 少个 n），
            // 只靠第一个词判断"像不像块头"，先接管、再让 UI 线程用真数据复核。
            Ok("header-prefix/半截 For", Editing.LooksLikeBlockHeaderPrefix("    For i = 1 To "));
            Ok("header-prefix/半截 Then", Editing.LooksLikeBlockHeaderPrefix("    If x > 0 The"));
            Ok("header-prefix/Select Case", Editing.LooksLikeBlockHeaderPrefix("    Select Case x"));
            Ok("header-prefix/#If", Editing.LooksLikeBlockHeaderPrefix("#If VBA7 Then"));
            Ok("header-prefix/Private Sub", Editing.LooksLikeBlockHeaderPrefix("Private Sub Foo("));
            Ok("header-prefix/整行注释", Editing.LooksLikeBlockHeaderPrefix("    ' 说明"));
            Ok("header-prefix/普通语句不算", !Editing.LooksLikeBlockHeaderPrefix("    x = 1"));
            Ok("header-prefix/字符串里的 For 不算", !Editing.LooksLikeBlockHeaderPrefix("    MsgBox \"For\""));
            Ok("header-prefix/续行符不算", !Editing.LooksLikeBlockHeaderPrefix("    For i = _"));
            Ok("header-prefix/空行不算", !Editing.LooksLikeBlockHeaderPrefix(""));

            // ---- 过程内跳转的两个落点Ctrl+↑ / Ctrl+↓）----
            // 行号全是 1-based：过程头 = 1，收尾行 = 最后一行。
            var nav = new List<string>
            {
                "Sub Foo()",             // 1
                "    Dim a As Long",     // 2
                "    x = 1",             // 3
                "    Dim b As Long",     // 4  <- 最后一个 Dim
                "    y = 2",             // 5
                "End Sub",               // 6
            };
            Eq("nav/↑ 落到最后一个 Dim 行", 4, Editing.LastDimEndLine(nav, 1, nav.Count));
            Eq("nav/↓ 落到收尾行上方最后一条代码", 5, Editing.ProcLastCodeLine(nav, 1, nav.Count));

            // 跨续行的 Dim：行尾要取【语句结束行】，不能停在 `Dim a, _` 上
            var navCont = new List<string>
            {
                "Sub Foo()",
                "    Dim a, _",
                "        b As Long",      // 3 <- 语句结束行
                "End Sub",
            };
            Eq("nav/↑ 跨续行取语句结束行", 3,
                Editing.LastDimEndLine(navCont, 1, navCont.Count));

            // Const / Static 不是"Dim 行"：口径就是 dim 关键字
            var navConst = new List<string>
            {
                "Sub Foo()",
                "    Dim a As Long",     // 2 <- 最后一个（也是唯一的）Dim
                "    Const N = 3",
                "    Static s As Long",
                "End Sub",
            };
            Eq("nav/↑ Const·Static 不抢 Dim 的位置", 2,
                Editing.LastDimEndLine(navConst, 1, navConst.Count));

            // 一条 Dim 都没有 / 全是注释 => 0（调用方据此把键还回去）
            Eq("nav/↑ 没有 Dim 返回 0", 0, Editing.LastDimEndLine(
                new List<string> { "Sub Foo()", "    ' Dim x", "    x = 1", "End Sub" }, 1, 4));

            // 收尾行上方是空行：往上找最近一条代码，不能落在空行的行首
            var navBlank = new List<string>
            {
                "Sub Foo()",
                "    x = 1",             // 2
                "",
                "End Sub",
            };
            Eq("nav/↓ 跳过尾行上方的空行", 2,
                Editing.ProcLastCodeLine(navBlank, 1, navBlank.Count));

            // 空过程体：没有"上一行"可落
            Eq("nav/↓ 空过程体返回 0", 0, Editing.ProcLastCodeLine(
                new List<string> { "Sub Foo()", "End Sub" }, 1, 2));
            Eq("nav/↑ 空过程体返回 0", 0, Editing.LastDimEndLine(
                new List<string> { "Sub Foo()", "End Sub" }, 1, 2));

            // 只在【本过程】里找：下一个过程的 Dim 绝不能被捞过来
            var navTwo = new List<string>
            {
                "Sub A()",               // 1
                "    x = 1",
                "End Sub",               // 3
                "Sub B()",
                "    Dim z As Long",     // 5
                "End Sub",
            };
            Eq("nav/↑ 不跨过程捞 Dim", 0, Editing.LastDimEndLine(navTwo, 1, 3));
            Eq("nav/↑ 过程 B 自己的 Dim", 5, Editing.LastDimEndLine(navTwo, 4, 6));
        }

        // ------------------------------------------------------------------
        // 续行判据：`Ctrl+↑` 跨续行找最后的 Dim 时要用它，判错会把光标停在
        // 半条语句上（`Dim a, _` 那行）。删掉一键格式化（Ctrl+Shift+L）后，
        // 这是原 Indent.cs 里唯一还有调用方的成员，已并入 Editing。
        // ------------------------------------------------------------------
        private static void TestContinuation()
        {
            Ok("continuation/续行符", Editing.IsContinuation("x = 1 _"));
            Ok("continuation/下划线标识符不算续行", !Editing.IsContinuation("x = foo_"));
            Ok("continuation/短行不算续行", !Editing.IsContinuation("_"));
            Ok("continuation/Tab 打头也算", Editing.IsContinuation("x = 1\t_"));
            Ok("continuation/空串不崩",
                !Editing.IsContinuation("") && !Editing.IsContinuation(null));
        }

        // ------------------------------------------------------------------
        // 自动弹出的总判据**行号变了就不能弹**
        // 起因："有时候鼠标单击到某个变量名，它会自动弹提示词框"。
        // 单击另一行上的变量时，快照的"行号 + 行文本"全变了，和"敲了一个字"
        // 长得一模一样；拦不住它的正是"最近 1.2 秒按过键"那条（打完字顺手点一下）。
        // ------------------------------------------------------------------
        private static void TestAutoTrigger()
        {
            // 正常打字：同行、内容变了、刚按过键 -> 弹
            Ok("auto-trigger/同行打字弹",
                Rules.AutoTriggerWanted(true, false, true, false));
            // 鼠标单击另一行的变量（或方向键上下）：行号跳了 -> 不弹
            Ok("auto-trigger/行号跳了不弹",
                !Rules.AutoTriggerWanted(true, false, true, true));
            // 右键粘贴 / 拖拽 / 别的程序改模块：最近没按键 -> 不弹
            Ok("auto-trigger/没按键不弹",
                !Rules.AutoTriggerWanted(true, false, false, false));
            // 切了代码窗（哪怕刚按过键）-> 不弹
            Ok("auto-trigger/切模块不弹",
                !Rules.AutoTriggerWanted(true, true, true, false));
            // 什么都没变 -> 不弹
            Ok("auto-trigger/没变化不弹",
                !Rules.AutoTriggerWanted(false, false, true, false));

            // 剪贴板键不算"打字"：Ctrl+C/X/V 与 Shift+Insert。
            // 复制一个变量名粘到别处，光标停在词尾时会弹出候选窗 —— 根子就在这里：
            //   Ctrl+V 也是按键，"最近按过键"照样成立，于是被当成"敲完了一个词"。
            //   判定收在 Native.IsClipboardChord 一处。
            Ok("auto-trigger/Ctrl+V 算剪贴板键",
                Native.IsClipboardChord(true, false, false, Native.VK_V));
            Ok("auto-trigger/Ctrl+C 算剪贴板键",
                Native.IsClipboardChord(true, false, false, Native.VK_C));
            Ok("auto-trigger/Ctrl+X 算剪贴板键",
                Native.IsClipboardChord(true, false, false, Native.VK_X));
            Ok("auto-trigger/Shift+Insert 算剪贴板键",
                Native.IsClipboardChord(false, true, false, Native.VK_INSERT));
            Ok("auto-trigger/光按 V 不算剪贴板键",
                !Native.IsClipboardChord(false, false, false, Native.VK_V));
            Ok("auto-trigger/Alt+V 不算剪贴板键",
                !Native.IsClipboardChord(false, false, true, Native.VK_V));
            Ok("auto-trigger/Ctrl+Shift+V 仍算剪贴板键",
                Native.IsClipboardChord(true, true, false, Native.VK_V));

            // Ctrl+/ 的 `/` 键识别主键盘与小键盘除号都算，别的符号键不算。
            // "Ctrl+/ 失效"就出在这里 —— 钩子只认 VK_OEM_2，小键盘除号
            //   （VK_DIVIDE）按下去就像没发生一样。判定收在 Native.IsSlashKey 一处。
            Ok("auto-trigger/主键盘 / 算 slash 键", Native.IsSlashKey(Native.VK_OEM_2));
            Ok("auto-trigger/小键盘除号也算 slash 键", Native.IsSlashKey(Native.VK_DIVIDE));
            Ok("auto-trigger/别的符号键不算 slash", !Native.IsSlashKey(0xBC) && !Native.IsSlashKey('D'));
        }

        // ------------------------------------------------------------------
        // 开机自启："这条注册表值是不是我们这份 exe"
        // 只是纯判断，不读写注册表（真登记走 --startuptest）。
        //   值名必须完全对上：只按别的名字删的话，关了自启照样会被拉起来。
        // ------------------------------------------------------------------
        private static void TestAutoStart()
        {
            string me = AutoStart.ExePath();
            Ok("autostart/取得到自己的 exe 路径", !string.IsNullOrEmpty(me), me ?? "null");

            string dir = Path.GetDirectoryName(me);
            string name = Path.GetFileName(me);

            Ok("autostart/带引号的自己算我们的",
                AutoStart.IsOurCommand("\"" + me + "\""));
            Ok("autostart/带参数的自己算我们的",
                AutoStart.IsOurCommand("\"" + me + "\" --tray"));
            Ok("autostart/同目录同名（不带引号）也算",
                AutoStart.IsOurCommand(Path.Combine(dir, name)));

            // 别家的程序：文件名不同 -> 不是我们的，关闭自启时绝不许动它
            Ok("autostart/notepad 不是我们的",
                !AutoStart.IsOurCommand("\"C:\\Windows\\notepad.exe\""));
            Ok("autostart/空值不是我们的", !AutoStart.IsOurCommand(""));
            Ok("autostart/null 不是我们的", !AutoStart.IsOurCommand(null));

            // 同名但文件还在的"别家安装"不许当成我们的
            string other = Path.Combine("C:\\Windows", name);
            if (File.Exists(other))
                Ok("autostart/别处同名且文件还在 -> 不动它", !AutoStart.IsOurCommand(other));
            else
                Ok("autostart/别处同名但文件已失效 -> 算旧项", AutoStart.IsOurCommand(other));
        }

        // ------------------------------------------------------------------
        // 片段库：标题模糊匹配（不碰数据库，纯函数）
        // ------------------------------------------------------------------
        private static void TestSnippets()
        {
            List<int> hits;
            Ok("snippet/整标题相等", SnippetMatch.Score("getcode", "GetCode", out hits) != null);
            Ok("snippet/前缀", SnippetMatch.Score("Get", "GetCode", out hits) != null);
            Ok("snippet/子串", SnippetMatch.Score("Code", "GetCode", out hits) != null);
            Ok("snippet/驼峰子序列 gc -> GetCode", SnippetMatch.Score("gc", "GetCode", out hits) != null);
            Ok("snippet/不命中", SnippetMatch.Score("zzz", "GetCode", out hits) == null);
            Ok("snippet/多词 AND 全中", SnippetMatch.Score("get code", "GetCode", out hits) != null);
            Ok("snippet/多词 AND 有一个不中",
                SnippetMatch.Score("get zzz", "GetCode", out hits) == null);

            int s1 = SnippetMatch.Score("Get", "GetCode", out hits) ?? -1;
            int s2 = SnippetMatch.Score("Code", "GetCode", out hits) ?? -1;
            Ok("snippet/前缀分高于子串分", s1 > s2, "prefix=" + s1 + " substr=" + s2);

            var rows = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("Zeta", "z"),
                new KeyValuePair<string, string>("SetValue", "a"),
                new KeyValuePair<string, string>("Settings", "b"),
            };
            var f = SnippetMatch.Filter(rows, "");
            Eq("snippet/空词原样返回", f.Count, 3);
            Eq("snippet/空词保持原顺序", f[0].Title, "Zeta");

            var f2 = SnippetMatch.Filter(rows, "set");
            Ok("snippet/过滤掉不命中的", f2.Count == 2, "count=" + f2.Count);
            Ok("snippet/前缀命中排前面", f2.Count > 0 && f2[0].Title == "Settings",
                f2.Count > 0 ? f2[0].Title : "");
        }
    }
}
