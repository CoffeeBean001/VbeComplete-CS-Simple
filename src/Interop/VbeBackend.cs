// VBE 后端：读快照、读模块全文、把补全写回去、以及那些"改一行"的操作。
//
// 所有方法都只在 UI 线程调用，且【任何 COM 异常都被吞掉并返回失败】——
//   补全工具永远不允许把异常抛到用户脸上，更不允许因为 VBE 忙就把自己卡死。

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal sealed class Snapshot
    {
        public int Line;
        public string Text;
        public string Module;
        public int CaretCol;
        public override string ToString() => $"L{Line} C{CaretCol} {Module}";
    }

    /// <summary>
    /// 引擎需要的那一小撮能力。抽成接口只为一件事：让 Trigger / Accept 能在
    /// **不碰 COM** 的情况下被完整测试 —— 候选准不准全在这里，靠真机试一条
    /// bug 要试十次。
    /// </summary>
    internal interface IVbeSource
    {
        Snapshot Snapshot(bool verify = true);
        string ModuleText();
        List<string> ComponentNames();
        List<ModuleScope> ProjectScopes(out int signature);
        bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
            string completion, bool wantParens, bool wantSpace, string enclosingProc);
        string EnclosingProcName();
    }

    internal sealed class VbeBackend : IVbeSource
    {
        // 快照缓存：行列都没动就复用上一轮那份（读数从 8.7ms 降到 0.13ms 的关键）
        private Snapshot _snapCache;
        private DateTime _snapAt = DateTime.MinValue;
        private readonly bool _snapEnabled = !Settings.Flag("NO_SNAP_CACHE", false);
        private readonly double _snapMaxAge = Settings.Double("SNAP_MAX_AGE", 1.0, 0.05, 60);

        // 模块全文缓存
        private string _modText;
        private string _modName;
        private int _modCount = -1;
        private DateTime _modAt = DateTime.MinValue;

        // 工程组件缓存
        private List<string> _comps;
        private DateTime _compsAt = DateTime.MinValue;

        // 全工程作用域缓存：模块名 -> 分析结果。按"行数"判是否过期 ——
        // 行数没变就复用（读全文 + 跑正则比读一个 CountOfLines 贵两个数量级）。
        private readonly Dictionary<string, ModuleScope> _modScopes =
            new Dictionary<string, ModuleScope>(StringComparer.OrdinalIgnoreCase);
        private List<ModuleScope> _scopesList = new List<ModuleScope>();
        private static long _lastSelFailLog;    // GetSelection 失败日志的限流时刻
        private string _lastLineText;           // 上一次读到的那一行（判"行文本变了"）
        private string _projectKey;             // 当前工程身份（Name|FileName），换工程要清缓存
        private int _scopesSig = int.MinValue;
        private DateTime _scopesAt = DateTime.MinValue;
        private readonly double _scopesTtl = Settings.Double("SCOPE_TTL", 2.0, 0.2, 60);

        public static readonly bool AutoParen = !Settings.Flag("NO_AUTOPAREN", false);
        public static readonly bool AutoSpace = !Settings.Flag("NO_AUTOSPACE", false);

        public Action<char> SendChar;     // 由 main 注入（补空格用）
        public Action<int, bool, bool> SendVk;

        public void InvalidateSnapshot()
        {
            _snapCache = null;
            _modAt = DateTime.MinValue;
        }

        private void Fail(Exception e)
        {
            Log.Write("COM 失败：{0}", e.Message);
            ComCtx.Reset();
            _snapCache = null;
            _modName = null;
            _modCount = -1;
            _modAt = DateTime.MinValue;
            // 宿主重开 / 换了工程之后，缓存的 ModuleScope 属于**上一个工程**。
            //   默认模块名（Module1 / Sheet1）在哪个工程都一样，行数碰巧相同就会
            //   把别的项目的私有过程当成当前工程的。全部丢掉最干净。
            _modScopes.Clear();
            _scopesList = new List<ModuleScope>();
            _scopesSig = int.MinValue;
            _scopesAt = DateTime.MinValue;
        }

        /// <summary>
        /// 取"当前代码窗 + 它的 CodeModule"，拿不到返回 false。
        /// <para>为什么要有这一处：VBE 关着 / 切到没有代码窗的视图时 `ActiveCodePane`
        /// 是 null，直接点 `.CodeModule` 会抛 "runtime binding on a null reference"。
        /// 那**不是故障**，不该在日志里制造假警报，更不该让每个方法各判一遍
        /// （各写一遍的话必然漏判，统一收在这里）。</para>
        /// </summary>
        private bool TryCodePane(out dynamic pane, out dynamic module)
        {
            pane = null;
            module = null;
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                pane = vbe.ActiveCodePane;
                if (ReferenceEquals(pane, null)) return false;
                module = pane.CodeModule;
                return !ReferenceEquals(module, null);
            }
            catch (Exception e)
            {
                Fail(e);
                pane = null;
                module = null;
                return false;
            }
        }

        // ------------------------------------------------------------------
        // 读
        // ------------------------------------------------------------------

        /// <summary>
        /// 轻量快照：行号 / 行文本 / 模块名 / 光标列。多行选择或不在代码窗时返回 null。
        /// 两级读取：先只 GetSelection（0.13ms），行列没动就直接返回上一轮。
        /// </summary>
        public Snapshot Snapshot(bool verify = true)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe(verify);
                if (ReferenceEquals(vbe, null)) return null;
                dynamic cp = vbe.ActiveCodePane;
                if (ReferenceEquals(cp, null)) { _snapCache = null; return null; }

                int sl, sc, el, ec;
                if (!ComCall.CallFourOut(cp, "GetSelection", out sl, out sc, out el, out ec))
                {
                    // 这里不能一声不响地返回 null —— 上层会理解成"现在没有
                    //   代码窗"，于是"补全突然什么都不做了"时日志里一个字都没有。
                    //   轮询 50ms 一轮，不能每次都写，5 秒最多记一条。
                    long nowMs = Environment.TickCount64;
                    if (nowMs - _lastSelFailLog > 5000)
                    {
                        _lastSelFailLog = nowMs;
                        Log.Write("GetSelection 失败：拿不到选区（VBE 忙 / 无活动代码窗 / 宿主在切工程？）");
                    }
                    _snapCache = null;
                    return null;
                }
                if (sl != el) { _snapCache = null; return null; }

                if (_snapEnabled && _snapCache != null &&
                    _snapCache.Line == sl && _snapCache.CaretCol == ec &&
                    (DateTime.UtcNow - _snapAt).TotalSeconds <= _snapMaxAge)
                {
                    // 行列相同 ≠ 还是同一个模块。切到另一个模块后正好停在
                    //   同一行同一列（都停在最后一行很常见），直接返回缓存就会把
                    //   上个模块的行文本当成这一行 —— 候选池、取词、过程判定全错位。
                    //   多读一次 `CodeModule.Name`（与 GetSelection 同量级，0.1ms）
                    //   换掉这个隐患，比事后 Debug 便宜。
                    dynamic cmc = cp.CodeModule;
                    string mnow = "";
                    try { mnow = Convert.ToString(cmc.Name) ?? ""; } catch { }
                    if (string.Equals(mnow, _snapCache.Module, StringComparison.Ordinal))
                        return _snapCache;
                }

                dynamic cm = cp.CodeModule;
                string mod = "";
                try { mod = Convert.ToString(cm.Name) ?? ""; } catch { }
                string text = Convert.ToString(cm.Lines(sl, 1)) ?? "";

                var s = new Snapshot { Line = sl, Text = text, Module = mod, CaretCol = ec };
                if (_snapEnabled) { _snapCache = s; _snapAt = DateTime.UtcNow; }

                // 当前行文本变了 ⇒ 模块全文缓存作废（不只是光标动）
                if (mod != _modName)
                {
                    // 离开的这个模块可能刚被改过（改声明名、注释掉一行 ——
                    //   行数不变）。把它从作用域缓存里摘掉，下次 ProjectScopes
                    //   重算时会重新解析它；否则它的旧声明会一直挂在候选池里，
                    //   而新名字永远出不来（缓存判据只看行数）。
                    if (!string.IsNullOrEmpty(_modName)) _modScopes.Remove(_modName);
                    _modName = mod;
                    _modCount = -1;
                    _modAt = DateTime.MinValue;
                }
                // **行文本变了也要作废模块全文缓存**。上面那句注释一直是这么
                //   写的，但代码只判了模块名 —— 结果是 1.5 秒内刚写的 `Dim foo`
                //   看不见：`ModuleText()` 直接返回旧全文，池子里根本没有 foo，
                //   用户会以为"新变量不提示"。判据放在"真的读了当前行"这条路径上
                //   （缓存命中时不重读行，不额外付 COM）。
                else if (_modText != null && !string.IsNullOrEmpty(_lastLineText) &&
                         !string.Equals(text, _lastLineText, StringComparison.Ordinal))
                {
                    _modCount = -1;
                    _modAt = DateTime.MinValue;
                }
                _lastLineText = text;
                return s;
            }
            catch (Exception e) { Fail(e); return null; }
        }

        /// <summary>当前模块全文。带缓存：模块名 + 行数都没变就不重读。</summary>
        public string ModuleText()
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return null;
                dynamic cp = vbe.ActiveCodePane;
                if (ReferenceEquals(cp, null)) return null;
                dynamic cm = cp.CodeModule;
                string name = Convert.ToString(cm.Name) ?? "";
                int count = Convert.ToInt32(cm.CountOfLines);
                if (_modText != null && _modName == name && _modCount == count &&
                    (DateTime.UtcNow - _modAt).TotalSeconds < 1.5)
                    return _modText;

                string all = count > 0 ? Convert.ToString(cm.Lines(1, count)) : "";
                _modText = all ?? "";
                _modName = name;
                _modCount = count;
                _modAt = DateTime.UtcNow;
                return _modText;
            }
            catch (Exception e) { Fail(e); return null; }
        }

        /// <summary>"文本被我们改过"之后必须调它，否则下一轮读到的还是改动前的内容。</summary>
        public void OnWeWrote()
        {
            _modCount = -1;
            _modAt = DateTime.MinValue;
            InvalidateSnapshot();
            // 当前模块刚被我们改过：作用域缓存里那一条必须作废。否则"删掉一个
            // 声明、行数却没变"的情况下，候选池里还留着已经不存在的名字。
            if (!string.IsNullOrEmpty(_modName)) _modScopes.Remove(_modName);
            _scopesAt = DateTime.MinValue;
        }

        /// <summary>"纯文本里失效"的标记：确认模块缓存已经跟不上时强制重读一次。</summary>
        public void ForceReloadModule()
        {
            _modCount = -1;
            _modAt = DateTime.MinValue;
        }


        /// <summary>本工程所有组件的名字（模块 / 类模块 / 窗体）。缓存 30 秒。</summary>
        public List<string> ComponentNames()
        {
            if (_comps != null && (DateTime.UtcNow - _compsAt).TotalSeconds < 30) return _comps;
            var list = new List<string>();
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return list;
                dynamic proj = vbe.ActiveVBProject;
                if (ReferenceEquals(proj, null)) return list;
                dynamic comps = proj.VBComponents;
                int n = Convert.ToInt32(comps.Count);
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        dynamic c = comps.Item(i);
                        string nm = Convert.ToString(c.Name);
                        if (!string.IsNullOrEmpty(nm)) list.Add(nm);
                    }
                    catch { }
                }
                _comps = list;
                _compsAt = DateTime.UtcNow;
            }
            catch (Exception e) { Fail(e); }
            return list;
        }

        /// <summary>
        /// 当前工程的身份：`Name|FileName`。取不到返回空串（调用方据此跳过校验）。
        /// <para>FileName 对"从没保存过的工作簿"会抛 COM 异常 —— 那种情况只用 Name。
        /// 代价是两次 COM（Name + FileName），但只在作用域重算（默认 2 秒一次）时调。</para>
        /// </summary>
        private static string ProjectKey()
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return "";
                dynamic proj = vbe.ActiveVBProject;
                if (ReferenceEquals(proj, null)) return "";
                string name = Convert.ToString(proj.Name) ?? "";
                string file = "";
                try { file = Convert.ToString(proj.FileName) ?? ""; } catch { }
                if (name.Length == 0 && file.Length == 0) return "";
                return name + "|" + file;
            }
            catch { return ""; }
        }

        // ------------------------------------------------------------------
        // 读：全工程作用域
        // ------------------------------------------------------------------

        /// <summary>
        /// 全工程每个模块的作用域分析（含"别的模块的 Public 也能提示"）。
        ///
        /// 为什么必须有：候选池必须是全工程的，否则在 Module2 里写
        ///   `Public Sub Refresh()`、在 Module1 里敲 `ref` 什么都提示不出来 ——
        ///   而 VBA 里这是完全合法的跨模块调用，所以池子取全工程。
        ///
        /// signature 是"各组件名 + 行数"的滚动哈希，供引擎判"池子要不要重建"。
        /// 行数没变就复用上一次的分析结果（读全文 + 跑正则比读 CountOfLines
        /// 贵两个数量级，而打字时每秒要触发十几次）。
        /// </summary>
        public List<ModuleScope> ProjectScopes(out int signature)
        {
            // 先认工程身份。两个 .xlsm 的工程名都叫 `VBAProject`、组件也
            //   都叫 `Module1`/`Sheet1`，只按模块名做缓存键会把上本工作簿的
            //   Private 过程当成当前工程的（缓存键要用 `(Name, FileName)` +
            //   首尾组件名做隔离）。换工程就整体作废，宁可多解析一次。
            string pkey = ProjectKey();
            if (!string.IsNullOrEmpty(pkey) && !string.Equals(pkey, _projectKey, StringComparison.Ordinal))
            {
                _modScopes.Clear();
                _scopesList = new List<ModuleScope>();
                _scopesSig = int.MinValue;
                _scopesAt = DateTime.MinValue;
                _modName = null;
                _modCount = -1;
                _modAt = DateTime.MinValue;
                _comps = null;
                _projectKey = pkey;
            }

            if (_scopesSig != int.MinValue &&
                (DateTime.UtcNow - _scopesAt).TotalSeconds < _scopesTtl)
            {
                signature = _scopesSig;
                return _scopesList;
            }

            signature = _scopesSig;
            var list = new List<ModuleScope>();
            int sig = 17;
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return _scopesList;
                dynamic proj = vbe.ActiveVBProject;
                if (ReferenceEquals(proj, null)) return _scopesList;
                dynamic comps = proj.VBComponents;
                int n = Convert.ToInt32(comps.Count);

                var alive = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        dynamic comp = comps.Item(i);
                        string nm = Convert.ToString(comp.Name) ?? "";
                        if (nm.Length == 0) continue;
                        alive.Add(nm);

                        // 组件类型：1=标准模块 2=类模块 3=窗体 100=文档模块。
                        // 读不出来就当"未知"，池子那边按标准模块处理（宁可多提示，
                        // 也不要因为一次 COM 抖动把整个模块的名字全吞掉）。
                        int ctype = 0;
                        try { ctype = Convert.ToInt32(comp.Type); } catch { }

                        dynamic cm = comp.CodeModule;
                        int count = Convert.ToInt32(cm.CountOfLines);
                        sig = unchecked(sig * 31 + nm.ToLowerInvariant().GetHashCode() * 17 + count);

                        ModuleScope ms;
                        if (_modScopes.TryGetValue(nm, out ms) && ms.LineCount == count)
                        {
                            ms.ComponentType = ctype;   // 缓存命中也要刷新类型
                            list.Add(ms);
                            continue;
                        }
                        string text = count > 0 ? Convert.ToString(cm.Lines(1, count)) : "";
                        ms = Scope.Analyze(nm, text);
                        ms.ComponentType = ctype;
                        _modScopes[nm] = ms;
                        list.Add(ms);
                    }
                    catch { }   // 单个组件读不出来（被锁定 / 奇怪类型）不影响其它
                }

                // 已经删掉的模块，缓存也要跟着清
                if (_modScopes.Count > alive.Count)
                {
                    var dead = new List<string>();
                    foreach (var k in _modScopes.Keys) if (!alive.Contains(k)) dead.Add(k);
                    foreach (var k in dead) _modScopes.Remove(k);
                }

                _scopesList = list;
                _scopesSig = sig;
                _scopesAt = DateTime.UtcNow;
                signature = sig;
            }
            catch (Exception e)
            {
                Fail(e);
                // COM 抖到一半抛出来时，list 里只有前面几个模块的解析结果 ——
                //   直接返回它等于悄悄丢掉其余模块的名字（"某些名字不提示了"）。
                //   沿用上一次的完整结果，下一次轮询再整体重算。
                return _scopesList;
            }
            return list;
        }

        // ------------------------------------------------------------------
        // 写：补全
        // ------------------------------------------------------------------

        private static readonly Regex ReCallStmt =
            new Regex(@"^[ \t]*Call[ \t]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool InCallStatement(string line) =>
            line != null && ReCallStmt.IsMatch(line);

        /// <summary>
        /// 把 [wordStartCol, endCol) 换成 completion（都是 1-based 字符列）。
        /// 返回 true 表示真的写进去了。
        /// </summary>
        public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
            string completion, bool wantParens, bool wantSpace, string enclosingProc)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                dynamic cp = vbe.ActiveCodePane;
                if (ReferenceEquals(cp, null)) return false;
                dynamic cm = cp.CodeModule;

                string lineText = Convert.ToString(cm.Lines(lineNo, 1)) ?? "";
                string newLine = Rules.ReplaceWord(lineText, wordStartCol, endCol, completion);

                bool addedParens = false;
                if (wantParens && AutoParen && Rules.IsCallSite(lineText, wordStartCol)
                    && !InCallStatement(lineText))
                {
                    int k = wordStartCol - 1 + completion.Length;
                    if (k > newLine.Length) k = newLine.Length;
                    // 跳过空白【只用于判断】右边是不是已经紧跟左括号
                    int j = k;
                    while (j < newLine.Length && (newLine[j] == ' ' || newLine[j] == '\t')) j++;
                    if (!(j < newLine.Length && newLine[j] == '('))
                    {
                        newLine = newLine.Substring(0, k) + "()" + newLine.Substring(k);
                        addedParens = true;
                    }
                }

                int start0 = Math.Max(0, wordStartCol - 1);
                int end0 = start0 + completion.Length;

                cm.ReplaceLine(lineNo, newLine);
                OnWeWrote();

                // ReplaceLine 会重排整行（`,` 后、`:` 后、`=` 两边补空格），
                //   所以落光标的位置必须按【读回来的真行】算，不能拿写入前的下标。
                string actual = Convert.ToString(cm.Lines(lineNo, 1)) ?? "";
                if (actual != newLine)
                    end0 = Rules.FindCompletionEnd(actual, completion, end0 - completion.Length, end0);
                if (end0 > actual.Length) end0 = actual.Length;

                // `Rules.DispCol(text, i)` 给的是"字符 i【之前】"的 1-based 显示列
                //   （`_disp_width(...) + 1` 的同一语义）。
                //   * 不补括号：光标就在词尾 = 字符 end0 之前 ⇒ DispCol(end0)。
                //   * 补了括号：end0 指向那个 `(`，而我们要的是【两个括号中间】
                //     = 字符 end0+1（即 `)`）之前 ⇒ DispCol(end0 + 1)。
                //   写成了 DispCol(end0)，于是停在 `(`【左边】
                //     （`Split|()`），否则会"落在左括号左侧"。
                int caretCol = Rules.DispCol(actual, Math.Min(actual.Length, end0 + (addedParens ? 1 : 0)), 4);
                if (caretCol < 0) caretCol = 0;

                cp.SetSelection(lineNo, 1, lineNo, 1);
                cp.SetSelection(lineNo, caretCol, lineNo, caretCol);

                // 二次确认：极少数情况下 VBE 是在我们读完行【之后】
                // 才把括号补上的，那一对 `()` 会把光标挤到左括号左边。再读一次行，
                // 若它又变了，就按新行重新算一次（宁多重算一次，也不留错位）。
                if (addedParens)
                {
                    string actual2 = Convert.ToString(cm.Lines(lineNo, 1)) ?? "";
                    if (actual2 != actual)
                    {
                        int e2 = Rules.FindCompletionEnd(actual2, completion,
                            Math.Max(0, end0 - completion.Length), end0);
                        int c2 = Rules.DispCol(actual2, Math.Min(actual2.Length, e2 + 1), 4);
                        if (c2 != caretCol) cp.SetSelection(lineNo, c2, lineNo, c2);
                    }
                }

                bool wantSp = wantSpace && AutoSpace && !addedParens && !InCallStatement(lineText);
                if (wantSp && SendChar != null)
                {
                    // 空格不能靠改行文本补：VBE 会把行尾空格立刻 trim 掉
                    if (end0 >= actual.Length || actual[end0] != ' ')
                        SendChar(' ');
                }
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        // ------------------------------------------------------------------
        // 写：行操作
        // ------------------------------------------------------------------


        public bool ReplaceLineRaw(int lineNo, string text)
        {
            if (!TryCodePane(out _, out dynamic cm)) return false;
            try
            {
                cm.ReplaceLine(lineNo, text);
                OnWeWrote();
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public bool DeleteLines(int from, int count)
        {
            if (!TryCodePane(out _, out dynamic cm)) return false;
            try
            {
                cm.DeleteLines(from, count);
                OnWeWrote();
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public bool InsertLines(int at, string text)
        {
            if (!TryCodePane(out _, out dynamic cm)) return false;
            try
            {
                cm.InsertLines(at, text);
                OnWeWrote();
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public int LineCount()
        {
            // VBE 关着的时候 `ActiveCodePane` 是 null —— 那是"现在没有代码窗"，
            //   不是故障，安静返回 -1 就行（判空统一在 TryCodePane 里）。
            if (!TryCodePane(out _, out dynamic cm)) return -1;
            try { return Convert.ToInt32(cm.CountOfLines); }
            catch (Exception e) { Fail(e); return -1; }
        }

        public string LineAt(int lineNo)
        {
            if (!TryCodePane(out _, out dynamic cm)) return null;
            try { return Convert.ToString(cm.Lines(lineNo, 1)) ?? ""; }
            catch (Exception e) { Fail(e); return null; }
        }

        /// <summary>
        /// 一次读回连续 <paramref name="count"/> 行（VBE 用 CRLF 拼好）。
        /// Ctrl+/ 改走"一次读整段"——逐行 `Lines(i,1)` 是 count 次跨进程往返，
        ///   选 30 行就白白多付 29 次。读失败返回 null，调用方按"拿不到"处理。
        /// </summary>
        public string LinesBlock(int from, int count)
        {
            if (from < 1 || count < 1) return null;
            if (!TryCodePane(out _, out dynamic cm)) return null;
            try { return Convert.ToString(cm.Lines(from, count)); }
            catch (Exception e) { Fail(e); return null; }
        }

        // ==================================================================
        // VBE 内置命令 —— "设置注释块"/"解除注释块"
        //
        // 这两个 ID 是 VBE 自己的命令控件 ID，**与 VBE 的语言版本无关**
        //   （中文版叫"设置注释块/解除注释块"，英文版叫 Comment Block/Uncomment
        //   Block，ID 都是 192 / 2552）。所以绝不要按 Caption 找控件 ——
        //   按名字找在别的语言版本上必然找不到。
        //
        // 真机实测（Excel 365 / VBE 7.1）：
        //   · Execute() 一次 ≈ 3.5ms，我们逐行 ReplaceLine 三行 ≈ 15ms
        //   · 执行后**选区原样保留**，不用再 SetSelection 摆回去
        //   · 改动进 VBE 自己的 Undo 栈，一次 Ctrl+Z 就整块回来
        //   · FindControl 首次 18ms、之后 <1ms —— 必须缓存，否则每次按都白付 18ms
        // ==================================================================
        public const int CmdCommentBlock = 192;
        public const int CmdUncommentBlock = 2552;
        private const int MsoControlButton = 1;

        private object _cmdComment;     // 缓存住的 CommandBarControl（只在 UI 线程用）
        private object _cmdUncomment;

        /// <summary>
        /// 直接执行 VBE 自己的命令。返回 false 表示"没做成"——调用方**必须**回退到
        /// 手写路径，绝不能拿 true 当"大概成了"（吞键却什么都没发生是最坏的结果）。
        /// </summary>
        public bool RunVbeCommand(int id)
        {
            for (int attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    object ctl = id == CmdCommentBlock ? _cmdComment : _cmdUncomment;
                    if (ctl == null)
                    {
                        dynamic vbe = ComCtx.Vbe();
                        if (ReferenceEquals(vbe, null)) return false;
                        dynamic bars = vbe.CommandBars;
                        if (ReferenceEquals(bars, null)) return false;
                        // FindControl(Type, Id) —— 后面三个参数是可选的。
                        ctl = bars.FindControl(MsoControlButton, id);
                        if (ReferenceEquals(ctl, null)) return false;
                        if (id == CmdCommentBlock) _cmdComment = ctl; else _cmdUncomment = ctl;
                    }

                    dynamic d = ctl;
                    if (!d.Enabled)
                    {
                        // 命令控件存在、只是**当前上下文**不让用（VBE 不在前台 / 焦点
                        // 不在代码窗时亲眼见过 Enabled=false）。缓存本身是好的，别丢 ——
                        // 丢了下次还得再付一次 18ms 的 FindControl。
                        // 交给调用方回退手写路径，两边结果必须一模一样。
                        return false;
                    }
                    d.Execute();
                    return true;
                }
                catch (Exception e)
                {
                    // 绝大多数是"VBE 关掉又打开、缓存的 IDispatch 已失效"。
                    // 清缓存重取一次，第二次仍然失败才认输（并记日志）。
                    DropCmdCache(id);
                    if (attempt == 1) { Fail(e); return false; }
                }
            }
            return false;
        }

        private void DropCmdCache(int id)
        {
            if (id == CmdCommentBlock) _cmdComment = null;
            else if (id == CmdUncommentBlock) _cmdUncomment = null;
        }

        /// <summary>
        /// 把 text 写回第 lineNo 行。VBE 允许光标停在【模块最后一行的下一行】
        /// （一个尚不存在的"虚拟空行"）：那时 `Lines()` 读出来是空串、看着一切
        /// 正常，但 `ReplaceLine()` 会抛"无效的过程调用或参数"（真机
        /// 实测）。这种行得先用 `InsertLines()` 真的建出来。
        /// </summary>
        public bool WriteLineRaw(int lineNo, string text)
        {
            int n = LineCount();
            if (n >= 0 && lineNo > n)
            {
                if (lineNo != n + 1) return false;   // 差得离谱，不猜
                return InsertLines(lineNo, text ?? "");
            }
            return ReplaceLineRaw(lineNo, text);
        }

        // ------------------------------------------------------------------
        // 写：编辑手感（Shift+Enter / 回车缩进 / 退格 / 自动配对 / 片段）
        // ------------------------------------------------------------------

        /// <summary>本模块第 lineNo 行【上方】各行（越靠近它的越靠后）。</summary>
        private List<string> LinesAbove(string[] all, int lineNo)
        {
            var list = new List<string>();
            int last = Math.Min(lineNo - 2, all.Length - 1);
            for (int i = 0; i <= last; i++) list.Add(all[i]);
            return list;
        }

        /// <summary>本模块第 lineNo 行【下方】各行（越靠近它的越靠前）。</summary>
        private List<string> LinesBelow(string[] all, int lineNo)
        {
            var list = new List<string>();
            for (int i = lineNo; i < all.Length; i++) list.Add(all[i]);
            return list;
        }

        /// <summary>
        /// Shift+Enter（smart=false）/ 回车（smart=true）：在光标所在行【下方】
        /// 新起一行，光标落在缩进之后，本行一个字符都不动。
        ///
        /// smart=true 时多四道保险（任一不满足就返回 false，让调用方把回车原样
        /// 还给系统）：有选区、光标不在行尾（那是拆行）、空行、引号没闭合。
        /// smart=false 时刻意**不**要求光标在行尾 —— 那条路本来就不拆行。
        ///
        /// align_branch：本行是 Else / ElseIf / Case 时先把【本行】拉回该在的层级。
        /// auto_close：块头后面补一行收尾（保守口径见 Editing.CloserNeeded）。
        /// </summary>
        public bool NewLineBelow(bool smart, bool autoClose, bool alignBranch)
        {
            try
            {
                int sl, sc, el, ec;
                if (!TryGetSelection(out sl, out sc, out el, out ec)) return false;
                if (smart && (sl != el || sc != ec)) return false;      // 有选区：不接管
                int anchor = smart ? sl : Math.Max(sl, el);
                if (anchor < 1) return false;

                // 光标可能停在"模块末尾的虚拟空行"（CountOfLines+1）：那一行不存在，
                // 往下插会越界。夹回最后一行 —— 结果就是"在末尾追加一行"。
                int lineCount = LineCount();
                if (lineCount < 0) return false;
                if (anchor > lineCount) anchor = Math.Max(1, lineCount);

                // 必须重读：ModuleText 有 1.5 秒缓存，而这里要拿"下方有哪些行"
                //   去判"该不该补收尾" —— 拿旧文本会补出双份 `End If`（编译不过）。
                ForceReloadModule();
                string text = ModuleText();
                if (text == null) return false;
                var all = Parser.SplitLines(text);

                // anchor 可能指向"模块末尾的虚拟空行"（Lines 读出来是空串）
                string lineText = LineAt(anchor);
                if (lineText == null) lineText = "";

                // 真正的"要不要接管"在这里用**真 COM 数据**再判一次。
                //   钩子那一侧可能因为快照落后而"先接管了再说"（见 App 的回车分支），
                //   所以这里必须自己兜住：不是块头 / 整行注释、光标不在行尾、行内
                //   有 Tab —— 一律 false，调用方会把回车原样还给 VBE。
                if (smart && !Editing.EnterIndentWanted(lineText, ec)) return false;

                string closerText = null;
                string fixedLine = null;
                string indent;

                var above = LinesAbove(all, anchor);
                var forUnit = new List<string>(above);
                forUnit.Add(lineText);
                string unit = Editing.IndentUnitFor(forUnit);

                if (smart)
                {
                    if (lineText.Trim().Length == 0) return false;      // 空行：交给 VBE
                    if (Editing.UnterminatedString(lineText)) return false;  // 让 VBE 补右引号
                    int caretChar = Rules.CharColFromDisp(lineText, ec);
                    if (caretChar - 1 < lineText.Length &&
                        lineText.Substring(caretChar - 1).Trim().Length > 0)
                        return false;                                   // 光标后还有内容 = 拆行

                    indent = Editing.NextLineIndent(lineText, above, unit);
                    if (alignBranch)
                    {
                        string mine, next;
                        if (Editing.BranchAlign(lineText, above, unit, out mine, out next))
                        {
                            fixedLine = mine;
                            indent = next;
                        }
                    }
                    if (autoClose) closerText = CloserFor(lineText, above, all, anchor);
                }
                else
                {
                    if (lineText.Trim().Length == 0)
                    {
                        // 空行：不替用户猜层级，照抄本行空白（= VBE 原生的结果）
                        indent = Editing.LeadingWs(lineText);
                    }
                    else
                    {
                        indent = Editing.NextLineIndent(lineText, above, unit);
                        if (alignBranch)
                        {
                            string mine, next;
                            if (Editing.BranchAlign(lineText, above, unit, out mine, out next))
                            {
                                fixedLine = mine;
                                indent = next;
                            }
                        }
                    }
                    if (autoClose) closerText = CloserFor(lineText, above, all, anchor);
                }

                // 分支行先拉回它该在的层级：只动行首空白，行内一个字符都不碰
                if (fixedLine != null)
                {
                    string ws = Editing.LeadingWs(lineText);
                    if (ws != fixedLine)
                    {
                        string body = lineText.Substring(ws.Length);
                        if (!WriteLineRaw(anchor, fixedLine + body)) return false;
                    }
                }

                if (!InsertLines(anchor + 1, indent)) return false;
                if (closerText != null)
                {
                    if (!InsertLines(anchor + 2, closerText))
                    {
                        // 收尾插失败：把刚插的空行收掉，回到"什么都没发生"
                        try { DeleteLines(anchor + 1, 1); } catch { }
                        return false;
                    }
                }

                // 新行已经插进去了 —— 这一段再出异常也不能返回 false（否则调用方
                // 会再补一个回车，用户按一下就多出一行）。所以 SetCaret 必须单独
                // try：它的 COM 调用是会抛的（行号越界 / 模块被关掉 / RPC 断了），
                // 让它冒到下面那个 catch 里去就等于"吞了回车又补了一行"。
                // 列要按**显示列**算：SetCaret 吃的是 VBE 的显示列，一个 Tab 占
                //   好几个格。用 indent.Length（字符数）在 Tab 工程里会把光标落进
                //   缩进中间 —— 第一个字符就打在错的位置上。空格缩进两者相等。
                try { SetCaret(anchor + 1, Rules.CaretColForCharCol(indent, indent.Length + 1)); }
                catch (Exception e) { Log.Write("回车缩进：已插行、落光标失败（按成功返回）：{0}", e.Message); }
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>块头该补的收尾（含缩进）；不该补返回 null。</summary>
        private string CloserFor(string lineText, List<string> above, string[] all, int anchor)
        {
            string closer = Editing.BlockCloser(lineText);
            if (closer == null) return null;
            string baseWs = Editing.LeadingWs(lineText);
            var below = LinesBelow(all, anchor);
            // 必须把 kind 与"上方各行"一起传：判据要靠 kind 分辨"下面那个同类收尾
            //   是本块的还是外层块的/用户写歪的"（见 Editing.CloserNeeded 的说明）。
            if (!Editing.CloserNeeded(below, closer, Editing.IndentWidth(baseWs),
                    Editing.BlockKind(lineText), above)) return null;
            return baseWs + closer;
        }

        /// <summary>
        /// 退格删整行：光标所在行只有空白时一次删掉，光标落到【上一行行尾】
        /// （与 VBE 原生"空行上按退格"的结果一致）。
        /// </summary>
        public bool DeleteBlankLineHere()
        {
            try
            {
                int sl, sc, el, ec;
                if (!TryGetSelection(out sl, out sc, out el, out ec)) return false;
                if (sl != el || sc != ec) return false;         // 有选区：那是"删选区"

                int count = LineCount();
                if (count <= 0 || sl < 1 || sl > count) return false;

                string raw = LineAt(sl) ?? "";
                // 权威复核：喂进去的是【此刻真读到的】行与列 —— 快照最多 100ms 前的
                if (!Editing.BlankLineDeleteWanted(raw, ec)) return false;

                if (!DeleteLines(sl, 1)) return false;

                try
                {
                    if (sl <= 1) SetCaret(1, 1);
                    else CaretToLineEnd(sl - 1);
                }
                catch (Exception e) { Log.Write("删空行：行已删、落光标失败（按成功返回）：{0}", e.Message); }
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>
        /// 退格删一对：光标夹在一对**空**括号 / 空引号中间时，一次把两端都删掉，
        /// 光标停在左符号原来的位置。
        /// </summary>
        public bool DeletePairHere()
        {
            try
            {
                int sl, sc, el, ec;
                if (!TryGetSelection(out sl, out sc, out el, out ec)) return false;
                if (sl != el || sc != ec) return false;         // 有选区：那是"删选区"

                string raw = LineAt(sl);
                if (raw == null) return false;

                int left, right;
                if (!Editing.PairDeleteSpan(raw, ec, out left, out right)) return false;

                string newLine = raw.Remove(left, right - left + 1);
                if (!WriteLineRaw(sl, newLine)) return false;

                // VBE 可能把整行重排过：读回来对不上就按"压平前缀"退回，绝不越界
                string actual = LineAt(sl) ?? newLine;
                int caretOff = left;
                if (!string.Equals(actual, newLine, StringComparison.Ordinal))
                {
                    int off = LocateByPrefix(actual, newLine, left);
                    if (off >= 0) caretOff = off;
                }
                caretOff = Math.Max(0, Math.Min(caretOff, actual.Length));
                try { SetCaret(sl, caretOff + 1); }
                catch (Exception e) { Log.Write("删一对：已删、落光标失败（按成功返回）：{0}", e.Message); }
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>
        /// 在 VBE 重排过的行里找回我们那个落点：拿"压平后的前缀"当指纹 ——
        /// 取写入内容在 off0 之前的部分，去掉所有空白转小写，再在实际行上前缀匹配
        /// 同样长度的字符。找不到返回 -1（调用方退回原 off0）。
        /// </summary>
        private static int LocateByPrefix(string actual, string written, int off0)
        {
            if (string.IsNullOrEmpty(actual) || string.IsNullOrEmpty(written) || off0 < 0) return -1;
            if (off0 > written.Length) off0 = written.Length;
            string want = NormCode(written.Substring(0, off0));
            if (want.Length == 0) return -1;

            int i = 0;
            for (; i < actual.Length && want.Length > 0; i++)
            {
                char c = char.ToLowerInvariant(actual[i]);
                if (c == ' ' || c == '\t') continue;
                if (c != want[0]) return -1;
                want = want.Substring(1);
            }
            return want.Length == 0 ? i : -1;
        }

        private static string NormCode(string s)
        {
            var sb = new StringBuilder(s == null ? 0 : s.Length);
            foreach (char c in s ?? "")
            {
                if (c == ' ' || c == '\t') continue;
                sb.Append(char.ToLowerInvariant(c));
            }
            return sb.ToString();
        }

        /// <summary>
        /// 自动配对：敲 `(` / `"` 时替用户补出右半边，光标落在中间。
        /// 返回 true = "这一下我代为输入了"（调用方据此吞掉原按键）；
        /// 返回 false = "我没管"（调用方必须放行，绝不吞掉按键却什么都没发生）。
        ///
        /// openCh 可能是**全角**（中文标点输入法）—— 由钩子侧换算后送进来，
        /// 这一层绝不自己再去猜输入法状态（两份判据早晚会打架）。
        /// </summary>
        public bool InsertPair(char openCh)
        {
            char closeCh = Editing.CloseOf(openCh);
            if (closeCh == '\0') return false;

            bool wrote = false;
            try
            {
                int sl, sc, el, ec;
                if (!TryGetSelection(out sl, out sc, out el, out ec)) return false;

                if (sl != el || sc != ec)
                {
                    // 单行选区 + 开符号 -> 用符号把选中的内容包起来（`abc` -> `"abc"`）
                    bool wrap = openCh == '"' || openCh == '“' || openCh == '(' || openCh == '（';
                    if (sl == el && sc < ec && wrap)
                        return WrapSelection(sl, sc, ec, openCh, closeCh);
                    return false;     // 其余选区：替换语义交给 VBE，别插
                }

                string lineText = LineAt(sl);
                if (lineText == null) return false;

                int charCol = Rules.CharColFromDisp(lineText, ec);
                // 只挡"注释里"，不挡"字符串里"：光标停在未闭合字符串的末尾时，
                //   这一下引号正是收尾的那一半，必须由 PairInsertion 去补一个；
                //   停在已有右引号之前时还要"跨过去"（见 Rules.InCommentOnly）。
                if (Rules.InCommentOnly(lineText, charCol)) return false;
                if (InCallStatement(lineText)) return false;   // `Call foo` 里的空括号 VBE 会删掉

                int col0 = Math.Max(0, Math.Min(charCol - 1, lineText.Length));
                string newLine;
                int caretOff;
                bool paired;

                if (openCh == ')' || openCh == '）')
                {
                    // 右半边已经在光标右边 -> 只跨过去；否则放行，让 VBE 真插一个
                    if (col0 >= lineText.Length || lineText[col0] != openCh) return false;
                    newLine = null;
                    caretOff = col0 + 1;
                    paired = false;
                }
                else
                {
                    newLine = Editing.PairInsertion(lineText, col0, openCh, closeCh,
                        out caretOff, out paired);
                }

                if (newLine != null)
                {
                    if (!WriteLineRaw(sl, newLine)) return false;
                    wrote = true;
                }

                string actual = newLine == null ? lineText : (LineAt(sl) ?? newLine);
                if (newLine != null && !string.Equals(actual, newLine, StringComparison.Ordinal))
                {
                    // VBE 重排过整行：重新定位我们插入的那对符号
                    int off = paired
                        ? LocateInserted(actual, newLine, openCh, closeCh, col0)
                        : LocateByPrefix(actual, newLine, col0 + 1);
                    if (off >= 0) caretOff = off;
                    else caretOff = Math.Min(caretOff, actual.Length);
                }
                caretOff = Math.Max(0, Math.Min(caretOff, actual.Length));

                // 文本已经写进去了 —— 落光标失败不能把结果拉成 false，否则调用方
                // 会再注入一个开符号，行里就多出一个 `(`
                try { SetCaret(sl, caretOff + 1); }
                catch (Exception e) { Log.Write("配对：已写、落光标失败（按成功返回）：{0}", e.Message); }
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
            finally
            {
                if (wrote) OnWeWrote();
            }
        }

        /// <summary>在重排过的行里找 `openCh+closeCh`，取离 col0 最近的那一处（返回其"之后"的下标）。</summary>
        private static int LocateInserted(string actual, string written,
            char openCh, char closeCh, int col0)
        {
            if (string.IsNullOrEmpty(actual)) return -1;
            string pair = new string(new[] { openCh, closeCh });
            int best = -1;
            int i = 0;
            while (i <= actual.Length - pair.Length)
            {
                int hit = actual.IndexOf(pair, i, StringComparison.Ordinal);
                if (hit < 0) break;
                if (best < 0 || Math.Abs(hit - col0) < Math.Abs(best - col0)) best = hit;
                i = hit + 1;
            }
            if (best >= 0) return best + openCh.ToString().Length;
            return LocateByPrefix(actual, written, col0 + 1);
        }

        /// <summary>单行选区用一对符号包起来（`abc` -> `"abc"` / `(abc)`）。</summary>
        private bool WrapSelection(int sl, int sc, int ec, char openCh, char closeCh)
        {
            try
            {
                string lineText = LineAt(sl);
                if (lineText == null) return false;

                int a = Math.Max(0, Math.Min(Rules.CharColFromDisp(lineText, sc) - 1, lineText.Length));
                int b = Math.Max(a, Math.Min(Rules.CharColFromDisp(lineText, ec) - 1, lineText.Length));
                if (b <= a) return false;        // 实际没选中字符：交给 VBE

                string newLine = lineText.Substring(0, a) + openCh +
                                 lineText.Substring(a, b - a) + closeCh +
                                 lineText.Substring(b);
                if (!WriteLineRaw(sl, newLine)) return false;

                string actual = LineAt(sl) ?? newLine;
                int tgt = a + 1 + (b - a) + 1;
                if (!string.Equals(actual, newLine, StringComparison.Ordinal))
                {
                    int off = LocateByPrefix(actual, newLine, tgt);
                    if (off >= 0) tgt = off;
                }
                tgt = Math.Max(0, Math.Min(tgt, actual.Length));
                try { SetCaret(sl, tgt + 1); }
                catch (Exception e) { Log.Write("包裹选区：已写、落光标失败（按成功返回）：{0}", e.Message); }
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>
        /// 把代码窗光标上/下移 delta 行，尽量保持同一列。
        /// 不用"放行按键让 VBE 自己动"：VBE 里 Shift+↑/↓ 的原生语义是**扩选**。
        /// 列原样透传（VBE 的列与 SetSelection 的列是同一套语义），目标行短时由
        /// VBE 自己钳到行尾 —— 这里绝不做字符数换算。
        /// </summary>
        public bool MoveCaretLine(int delta)
        {
            try
            {
                int sl, sc, el, ec;
                if (!TryGetSelection(out sl, out sc, out el, out ec)) return false;
                int line, col;
                if (el >= sl) { line = el; col = ec; } else { line = sl; col = sc; }
                int target = line + delta;
                int count = LineCount();
                if (line < 1 || target < 1 || count < 0 || target > count) return false;

                // 先按目标行把列钳回它自己的行尾：`SetSelection` 给一个超过行长
                //   的列不是"宽容地钳住"，而是直接抛 COM 异常 —— 那样一来
                //   Fail() 会打满退避，接下来一两秒连候选都不弹了，用户看到的是
                //   "按了下 Shift+↑ 提示就卡住了"。钳一下就没有这个失败路径。
                string tgt = LineAt(target);
                if (tgt != null)
                {
                    int eol = Rules.DispCol(tgt, tgt.Length);
                    if (col > eol) col = eol;
                }
                if (col < 1) col = 1;

                if (!TryCodePane(out dynamic cp, out _)) return false;
                cp.SetSelection(target, col, target, col);
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>
        /// 把一段（可能多行的）代码片段插到光标处。落点模型是整行拼接：
        ///     前半行 + 片段第 1 行
        ///     片段第 2..n-1 行
        ///     片段最后 1 行 + 后半行
        /// 于是"光标在行中间"也能插；有选区则是替换语义（单行 / 多行都支持）。
        ///
        /// 写入"尽量原子"：中间 n-1 行用**一次** InsertLines 灌进去，再覆盖首行；
        /// 任何一步失败都回到"什么都没发生"（绝不留下半截片段 —— 那会让用户再
        /// 点一次，片段就插两遍）。
        /// </summary>
        public bool InsertSnippet(string text)
        {
            text = (text ?? "").Replace("\r\n", "\n").Replace('\r', '\n');
            if (text.Trim().Length == 0) return false;

            int sl = 0, k = 0;
            try
            {
                int sc, el, ec;
                if (!TryGetSelection(out sl, out sc, out el, out ec)) return false;
                if (sl < 1 || el < sl) return false;

                int count = LineCount();
                if (count < 0) return false;

                string first = sl > count ? "" : (LineAt(sl) ?? "");
                string last = sl == el ? first : (el > count ? "" : (LineAt(el) ?? ""));

                int i = Math.Max(0, Math.Min(Rules.CharColFromDisp(first, sc) - 1, first.Length));
                int colEnd = Rules.CharColFromDisp(last, ec);
                int j = sl == el
                    ? Math.Max(i, Math.Min(colEnd - 1, last.Length))
                    : Math.Max(0, Math.Min(colEnd - 1, last.Length));

                string head = first.Substring(0, i);
                string tail = last.Substring(j);

                var block = text.Split('\n');
                var output = new List<string>();
                if (block.Length == 1)
                {
                    output.Add(head + block[0] + tail);
                }
                else
                {
                    output.Add(head + block[0]);
                    for (int x = 1; x < block.Length - 1; x++) output.Add(block[x]);
                    output.Add(block[block.Length - 1] + tail);
                }
                k = output.Count - 1;

                // 光标可能停在模块末尾的虚拟空行：那一行还不存在，先建成真的空行
                bool virtualHead = sl > count && k > 0;
                if (virtualHead)
                {
                    if (!InsertLines(sl, "")) return false;
                }

                if (k > 0)
                {
                    var mid = new StringBuilder();
                    for (int x = 1; x < output.Count; x++)
                    {
                        if (x > 1) mid.Append("\r\n");
                        mid.Append(output[x]);
                    }
                    if (!InsertLines(sl + 1, mid.ToString()))
                    {
                        if (virtualHead) { try { DeleteLines(sl, 1); } catch { } }
                        Log.Write("片段：中间 {0} 行插入失败（原文未改动）", k);
                        return false;
                    }
                }

                if (!WriteLineRaw(sl, output[0]))
                {
                    // 半截片段比"失败"更糟：把刚插进去的行删掉
                    try
                    {
                        if (virtualHead) DeleteLines(sl, k + 1);
                        else if (k > 0) DeleteLines(sl + 1, k);
                    }
                    catch (Exception e) { Log.Write("片段：首行写入失败且回滚失败：{0}", e.Message); }
                    return false;
                }

                // 多行选区：把被替换掉的旧行收掉（插进来的 k 行把它们挤下去了）
                int drop = el - sl;
                if (drop > 0)
                {
                    try { DeleteLines(sl + 1 + k, drop); }
                    catch (Exception e) { Log.Write("片段：已插入，旧行删除失败：{0}", e.Message); }
                }

                OnWeWrote();

                // 摆光标：落在片段最后一个字符之后（= 尾部原文的起点）
                int lineNo = sl + k;
                string actual = LineAt(lineNo) ?? output[output.Count - 1];
                int pos = TailPos(actual, tail);
                try { SetCaret(lineNo, pos + 1); }
                catch (Exception e) { Log.Write("片段：已插入、落光标失败（按成功返回）：{0}", e.Message); }
                Log.Write("片段：插入 {0} 行 -> 行={1}", block.Length, lineNo);
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>尾部原文在写回后的那一行里从哪开始（找不到就退到行尾）。</summary>
        private static int TailPos(string actual, string tail)
        {
            if (string.IsNullOrEmpty(tail)) return actual == null ? 0 : actual.Length;
            if (string.IsNullOrEmpty(actual)) return 0;
            int p = actual.LastIndexOf(tail, StringComparison.Ordinal);
            return p >= 0 ? p : actual.Length;
        }

        /// <summary>把光标放到某行的行尾（列交给 VBE 自己钳）。</summary>
        public bool CaretToLineEnd(int lineNo)
        {
            if (!TryCodePane(out dynamic cp, out dynamic cm)) return false;
            try
            {
                int count = Convert.ToInt32(cm.CountOfLines);
                if (lineNo < 1 || lineNo > count) return false;
                string text = Convert.ToString(cm.Lines(lineNo, 1)) ?? "";
                int col = Rules.DispCol(text.TrimEnd('\r', '\n'), text.TrimEnd('\r', '\n').Length, 4);
                cp.SetSelection(lineNo, 1, lineNo, 1);
                cp.SetSelection(lineNo, col, lineNo, col);
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>
        /// 把光标放到第 charCol 个字符【之前】（1-based 字符列，与 VBE 的列语义一致）。
        /// 参数是"第几个字符之前"，不是"跳过几个字符"。直接把 charCol 喂给
        ///   DispCol，而 DispCol(i) 返回的是"字符 i 之前"的列 ⇒ 实际落在第 charCol+1 列，
        ///   每个调用点都稳定右偏一格（真机：SetCaret(10,1) 读回来是 (10,2)）。
        /// </summary>
        public bool SetCaret(int lineNo, int charCol)
        {
            if (!TryCodePane(out dynamic cp, out dynamic cm)) return false;
            try
            {
                string text = Convert.ToString(cm.Lines(lineNo, 1)) ?? "";
                int col = Rules.CaretColForCharCol(text, charCol, 4);
                cp.SetSelection(lineNo, 1, lineNo, 1);
                cp.SetSelection(lineNo, col, lineNo, col);
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>
        /// 设一个跨行选区（1-based 行 / VBE 的显示列）。
        /// 只有 `--hktest` 用它：格式化缩进"只理选区"的那条路，必须真的拉一个选区才算测到。
        /// </summary>
        public bool SetSelectionRange(int sl, int sc, int el, int ec)
        {
            try
            {
                if (!TryCodePane(out dynamic cp, out _)) return false;
                cp.SetSelection(sl, sc, el, ec);
                InvalidateSnapshot();
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>当前选区（四个 Long）。四个出参都是 ByRef，只能走 IDispatch 直调。</summary>
        public bool TryGetSelection(out int sl, out int sc, out int el, out int ec)
        {
            sl = sc = el = ec = 0;
            try
            {
                if (!TryCodePane(out dynamic cp, out _)) return false;
                return ComCall.CallFourOut((object)cp, "GetSelection", out sl, out sc, out el, out ec);
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public string SelectedText()
        {
            try
            {
                if (!TryCodePane(out dynamic cp, out dynamic cm)) return "";
                int sl, sc, el, ec;
                if (!ComCall.CallFourOut((object)cp, "GetSelection", out sl, out sc, out el, out ec)) return "";
                if (el < sl) return "";
                if (el == sl && ec <= sc) return "";
                return Convert.ToString(cm.Lines(sl, el - sl + 1)) ?? "";
            }
            catch (Exception e) { Fail(e); return ""; }
        }

        public bool FocusVbe()
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (!ReferenceEquals(vbe, null))
                {
                    try { vbe.MainWindow.SetFocus(); return true; } catch { }
                }
            }
            catch (Exception e) { Fail(e); }

            // COM 走不通时的兜底：直接按 VBE 主框架的类名把前台抢回来
            try
            {
                IntPtr frame = Native.FindWindow("wndclass_desked_gsk", null);
                if (frame != IntPtr.Zero) return Native.SetForegroundWindow(frame);
            }
            catch { }
            return false;
        }

        public string EnclosingProcName()
        {
            var snap = Snapshot();
            if (snap == null) return null;
            string text = ModuleText();
            if (text == null) return null;
            var procs = Parser.ExtractProcedures(text);
            var p = Parser.ProcOfLine(procs, snap.Line);
            return p?.Name;
        }
    }
}
