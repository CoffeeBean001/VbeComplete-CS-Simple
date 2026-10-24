// 候选引擎：攒池子、匹配、排序、决定弹不弹、以及"选完之后该补什么"。
//
// 这个类不引用 WinForms 的任何东西 —— 窗口通过回调接进来。
//   于是引擎的全部行为（弹不弹、排第几、补括号还是补空格）都能在没有界面的
//   情况下测。
//
// （口径）：池子里**只放自己写的名字** —— 过程名、变量名、参数名、
//   常量名。VBA 关键字 / 内建函数 / 内建类型 / 模块名 / 表名 / 窗体名 /
//   类名 / 自定义类型名一律不进候选（少弹几百条，剩下的全是能用的）。
//   `Reserved` 里那两张表只用来挡"没用 Dim 声明过的自由标识符"，不再出候选。
//
// 排序与匹配的口径来自真机逐条调试。
//   最关键的修正是【排序键的顺序】：
//     旧：作用域 > 命中位置 > 跨度 > 最近使用 > ...
//     新：命中质量 > 命中位置 > 跨度 > 作用域 > 最近使用 > 名字长度
//   旧顺序把"工程内的名字"一律顶到最前，于是一个 8 字母变量只是"松散命中"
//   也能排在精确命中的名字前面 —— 用户看到的就是"最像的那个不在第一个"。

using System;
using System.Collections.Generic;
using System.Linq;

namespace VbeComplete
{
    internal sealed class Candidate
    {
        public string Name;
        public string Kind;        // func / prop / var —— 图标档位
        public int Pos;            // 首个命中下标（越小越好）
        public int Span;           // 命中跨度（越小越好）
        public int[] Hits;         // 逐字符命中的下标（UI 高亮用；空 = 全都算命中）
        public int HitKind;        // 4 完全相等 / 3 前缀 / 2 词首缩略 / 1 连续块 / 0 分散
        public int ScopeKind;      // 4 本过程局部 / 3 本模块 / 2 别的模块 Public
        public int RecentRank;     // 最近使用次数
    }

    internal sealed class Engine
    {
        public IVbeSource Backend;
        public Recent Recent;

        // 一屏最多几行（15）。候选不足时
        //   窗口自己变矮 —— 口径是"显示几行就几行高"，不是固定 15 行留白。
        public readonly int ViewRows = Settings.Int("VIEW_ROWS", 15, 3, 30);
        public readonly bool HideOnComplete = Settings.Flag("HIDE_ON_COMPLETE", true);
        // 不是 readonly —— 探针要能当场切成 1 做 A/B（证明"点号后不弹"
        //   确实是这条判据在起作用），运行期只读一次 settings.ini。
        public bool MemberPopup = Settings.Flag("MEMBER_POPUP", false);
        public readonly bool DeclPopup = Settings.Flag("DECL_POPUP", false);

        // UI 回调
        public Action<IReadOnlyList<Candidate>, int, int> OnShow;   // items, selected, top
        public Action OnHide;
        public Action OnMove;                                        // 选中项变了，重画
        /// <summary>
        /// "现在该不该让位给 VBE 自己的成员列表"。非 null 且返回 true 时本次
        /// 触发直接收起 —— 两个列表叠在一起，用户分不清在看哪一个。
        /// 抽成回调是为了让引擎本身不依赖 Win32（自检里传 null 就行）。
        /// </summary>
        public Func<bool> YieldToVbe;

        private List<Candidate> _items = new List<Candidate>();
        private int _sel;
        private int _top;
        public bool Visible { get; private set; }

        // 池：全工程的声明。按行数缓存在 VbeBackend 那一层（ProjectScopes），
        //   这里只留当前这一轮的解析结果（池本身缓存在 VbeBackend.ProjectScopes）。
        private List<ModuleScope> _scopes = new List<ModuleScope>();
        private List<string> _poolNames = new List<string>();
        // 池里哪些名字【必须】通过作用域判定才能出现（= 声明出来的名字）。
        // 只有它们受作用域约束；"引用过但没声明"的隐式变量也是按作用域判的。
        private readonly HashSet<string> _poolScoped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _procKinds =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private string _poolKey;

        // 刚按 Tab 确认过的那个词。同一行、同一个词不再提示自己 ——
        // 口径："按完 Tab 意味着本次录入完成，不要再匹配提示"。
        private string _acceptedWord;
        private string _acceptedModule;
        private int _acceptedLine;
        private long _acceptedTicks;
        private const long AcceptQuietMs = 2000;

        // 这一次触发时"光标前的词是不是已经打全了"（打全就不再挂着列表）
        private bool _wordComplete;

        // 让位给 VBE 成员列表时的宽限状态（键 = 模块|行|显示列）。
        private string _yieldKey;
        private long _yieldUntil;
        private const int YieldGraceMs = 250;

        // 这一轮候选是"哪个模块、哪一行、从哪个字符列起、什么词"算出来的。
        //   AcceptName 拿它复核：候选算出来之后用户可能按了 Home / End，或者
        //   用鼠标点到了同一行的别处 —— 行列变了但我们手里还是旧候选，
        //   照旧写进去就是"把候选插到行首"或者"把另一个词截成两半"。
        //   为 null 表示"没有上下文"（直接调 AcceptName 的探针走这条路）。
        private string _ctxModule;
        private int _ctxLine;
        private int _ctxWordStart;

        public IReadOnlyList<Candidate> Items => _items;
        public int SelectedIndex => _sel;
        public int TopIndex => _top;

        /// <summary>
        /// 上一次触发里被**截断掉**的命中总数（0 = 没截断）。
        /// 池子可能几百条，列表只留前 240 条 —— 这个数给自检与日志用。
        /// </summary>
        public int TruncatedFrom { get; private set; }

        // ------------------------------------------------------------------
        // 池子
        // ------------------------------------------------------------------

        /// <summary>
        /// 重建池子。key 带"本模块文本长度 + 全工程签名 + 被挖掉的词"：
        /// 本模块一边打字一边重解析（保证自己新写的变量立刻能提示），
        /// 别的模块只在行数变了之后才重解析（不然每次按键都要重扫整个工程）。
        ///
        /// blank*(1-based) 是"正在输入的那个词"的位置 —— 解析前把它挖成空格，
        /// 候选池里就不会出现"用户正打到一半的那个名字自己"。
        /// </summary>
        private void EnsurePool(string moduleName, string moduleText,
            int blankLine = 0, int blankStart = 0, string blankToken = null)
        {
            int sig = 0;
            List<ModuleScope> others = null;
            if (Backend != null)
            {
                others = Backend.ProjectScopes(out sig);
            }

            // 缓存键带上内容哈希：只比长度的话，等长改名（`foo` -> `bar`）会被当成
            //   没变化，池子里继续留着旧名字。文本本来就已经全量取到手了，算一次
            //   哈希比再取一遍 COM 便宜得多。
            string key = (moduleName ?? "") + "|" +
                         (moduleText == null ? -1 : moduleText.Length) + ":" +
                         (moduleText == null ? 0 : moduleText.GetHashCode()) + "|" +
                         sig + "|" + blankLine + ":" + blankStart + ":" + (blankToken ?? "");
            if (_poolNames.Count > 0 && _poolKey == key) return;

            var scopes = new List<ModuleScope>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _procKinds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 1) 当前模块：用传进来的文本（最新鲜），不吃缓存；先把半截词挖掉
            ModuleScope cur = null;
            if (moduleText != null)
            {
                string parsed = !string.IsNullOrEmpty(blankToken)
                    ? Rules.BlankToken(moduleText, blankLine, blankStart, blankToken)
                    : moduleText;
                cur = Scope.Analyze(moduleName ?? "", parsed);
                scopes.Add(cur);
                seen.Add(moduleName ?? "");
            }

            // 2) 其他模块
            if (others != null)
                foreach (var ms in others)
                {
                    if (!seen.Add(ms.Module)) continue;
                    scopes.Add(ms);
                }
            _scopes = scopes;

            // 3) 合并名字
            //
            // 这里同时是"哪些名字不许进候选"的唯一闸门。口径：
            //   只要自己写的函数名 / 过程名 / 变量名，其余一律不弹。于是被挡掉的是
            //     * 组件名（模块名 / Sheet1 / ThisWorkbook / 窗体名 / 类名）——
            //       它们会被"无 Dim 的裸用"带进 FreeIdents（`Sheet1.Range(...)`），
            //       一并挡在这里，不用再去问一遍 COM；
            //     * 自定义类型 / 枚举名（`Type Person` / `Enum Color`）；
            //     * `Type` 的成员（只能 `p.X` 访问，绝不是裸名）；
            //     * VBA 关键字 / 内建函数 —— 它们压根不走这条路（见 Reserved）。
            var names = new List<string>();
            var nameSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _poolScoped.Clear();

            var componentNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var ms0 in scopes)
                if (!string.IsNullOrEmpty(ms0.Module)) componentNames.Add(ms0.Module);

            foreach (var ms in scopes)
            {
                // 对象模块（类模块 2 / 窗体 3 / 文档模块 100）的声明是**成员**，
                //   只能 `c.s1` 这样访问，绝不能当裸名提示。
                //   把它们并进池子，用户在别的模块里打一个 `s` 就会看到 `s1`
                //   —— 用户 报的"变量泄露"就是这条（cls1 里的 Public s1）。
                //   正在编辑的那个模块例外：在自己模块里它当然是裸名。
                //   类型未知（0，合成测试 / COM 抖动）按标准模块处理，宁可多提示。
                bool isCur = ReferenceEquals(ms, cur) ||
                             string.Equals(ms.Module, moduleName ?? "",
                                           StringComparison.OrdinalIgnoreCase);
                if (!isCur && ms.ComponentType != 0 && ms.ComponentType != 1) continue;

                foreach (var d in ms.Decls)
                {
                    if (d.KindHint == "type" || d.KindHint == "enum") continue;   // 类型名 / 枚举名不提示
                    if (d.Proc != null && d.Proc.StartsWith("type:", StringComparison.Ordinal))
                        continue;                                                  // Type 成员只能 p.X
                    if (componentNames.Contains(d.Name)) continue;                  // 与模块同名
                    if (nameSeen.Add(d.Name)) names.Add(d.Name);
                    // 声明出来的名字一律要过作用域判定：查不到档位 = 当前位置
                    //   看不见它（别的过程的局部 / 别的模块的 Private），必须剔除。
                    _poolScoped.Add(d.Name);
                    if (d.Proc == null && !string.IsNullOrEmpty(d.KindHint))
                        _procKinds[d.Name] = d.KindHint;
                }
                foreach (var f in ms.FreeIdents)
                {
                    if (componentNames.Contains(f)) continue;
                    if (nameSeen.Add(f)) names.Add(f);
                    // 无 Option Explicit 时"用过但没声明"的名字，在 VBA 里是
                    //   **模块私有**的隐式变量（别的模块看不见）。只把它并进池子、
                    //   不进 _poolScoped，于是 A 模块里随手写的 `foo = 1` 会让 B 模块
                    //   打 `fo` 也冒出 `foo` —— 用户眼里的"别的模块的变量跑我这儿来了"。
                    //   进 _poolScoped 后，当前模块靠 Visible() 的 rank 2 保留，
                    //   别的模块查不到档位就被剔除。
                    _poolScoped.Add(f);
                }
                foreach (var p in ms.Procs)
                {
                    if (nameSeen.Add(p.Name)) names.Add(p.Name);
                    _poolScoped.Add(p.Name);
                    _procKinds[p.Name] = p.Kind;
                }
            }

            _poolNames = names;
            _poolKey = key;
        }

        /// <summary>把每个名字折算成"作用域档位"，key 为小写名。</summary>
        private Dictionary<string, int> BuildRankMap(string curModule, string curProc)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var ms in _scopes)
            {
                bool isCur = string.Equals(ms.Module, curModule, StringComparison.OrdinalIgnoreCase);
                // 对象模块（类模块 2 / 窗体 3 / 文档模块 100）的声明是**成员**，
                //   只能 `c.s1` 这样访问 —— EnsurePool 跳过它们，这里也必须跳，
                //   于是类模块的 Public 成员仍然进了 rank 表。后果很隐蔽：只要有同名的
                //   名字从**别的路**（比如别的模块里的隐式变量 s1）进了池子，就会在
                //   这里被"认领"成 rank 2 并通过作用域判定 —— 用户看到的就是
                //   "模块1 里打 s 冒出了 cls1 的 s1"修过一次的那个泄露）。
                if (!isCur && ms.ComponentType != 0 && ms.ComponentType != 1) continue;
                var src = isCur ? ms.Visible(curProc) : ms.Exported();
                foreach (var kv in src)
                {
                    int r = kv.Value.Value;
                    int old;
                    if (map.TryGetValue(kv.Key, out old) && old >= r) continue;
                    map[kv.Key] = r;
                }
            }
            return map;
        }

        /// <summary>把池子按给定源码算出来（自检用；运行期由 Trigger 调）。</summary>
        public void PrimePool(string moduleText, string moduleName)
        {
            EnsurePool(moduleName ?? "", moduleText);
        }

        public void InvalidatePool()
        {
            _poolKey = null;
            _poolNames.Clear();
        }

        // ------------------------------------------------------------------
        // 触发
        // ------------------------------------------------------------------

        /// <summary>
        /// 判断现在该不该弹。manual=true 是 Ctrl+Space（豁免"打全即收"等收束规则）。
        /// 返回 true 表示窗口已显示。
        /// </summary>
        public bool Trigger(bool manual, Snapshot snap = null)
        {
            snap = snap ?? Backend?.Snapshot();
            if (snap == null) { Hide(); return false; }

            // VBE 自己的成员列表正画着 —— 一律让位。让位不是"不提示"：
            // 它的列表比我们的更准（有真类型信息），等它关掉之后下一次触发
            // 我们再补位。
            //
            // 让位带一个 250ms 的宽限期（对齐 PY 的 YIELD_GRACE）。VBE 的
            //   列表要等一个 WM_TIMER 才 Show，我们往往比它先跑完 —— 只判"此刻
            //   可见"的话，我们会在它前面闪一下再被压住。同一位置让过一次后，
            //   接下来 250ms 内继续让；过了就照常弹，不会"这里永远不提示"。
            if (!manual)
            {
                string ykey = (snap.Module ?? "") + "|" + snap.Line + "|" + snap.CaretCol;
                if (YieldToVbe != null && YieldToVbe())
                {
                    _yieldKey = ykey;
                    _yieldUntil = Environment.TickCount64 + YieldGraceMs;
                    Hide();
                    return false;
                }
                if (string.Equals(_yieldKey, ykey, StringComparison.Ordinal) &&
                    Environment.TickCount64 < _yieldUntil)
                {
                    Hide();
                    return false;
                }
            }

            string line = snap.Text ?? "";
            // GetSelection 给的是**显示列**，后面所有取词/判据都按字符列算
            //   —— 含中文/全角的行不换算就会取到空词，连弹窗都不出现。
            int caret = Rules.CharColFromDisp(line, snap.CaretCol);

            // 注释（`'` 或 `Rem`）里 / 字符串里一律不弹 —— 判据见 Rules.InDeadZone。
            // （对标 IDEA）：手动 Ctrl+Space 也拦。IDEA 在注释/字符串里同样不给
            //   代码补全 —— 那里弹出的候选按 Tab 只会把注释或字符串内容改写掉，
            //   没有任何语义价值。手动 Ctrl+Space 也不例外。
            if (Rules.InDeadZone(line, caret)) { Hide(); return false; }
            // 声明位置一律不弹（含 `As` 之后、裸声明 `Public x` 之后）。
            //   见 Rules.InDeclarationPosition 的说明。
            if (!manual && !DeclPopup && Rules.InDeclarationPosition(line, caret)) { Hide(); return false; }

            // 先取词：池子要拿"这个词被挖掉之后"的文本去解析。
            int wordStart;
            string word = Rules.WordBefore(line, caret, out wordStart);

            // `As` / `New` 之后一律不弹（用户 口径）。
            //   与上面的 InDeclarationPosition 互补：那条管 `Dim x As |` 这种
            //   声明开头的整行，这条管 `Set c = New |` / `dic.Add 1, New |` /
            //   `For i As |` / 过程参数 `a As |` —— 行首都不是声明关键字，
            //   否则它们会照常弹出全工程候选。
            if (!manual && !DeclPopup && Rules.AfterTypeKeyword(line, wordStart))
            {
                Hide();
                return false;
            }

            // 点号之后**一个候选都不出**，整个让给 VBE（硬规矩）："输入 student.，点号后面就
            //   不要出提示框了"——我们的池子是裸名池（变量名 / 过程名 / 关键字），
            //   出在点号后面就是噪音；VBE 在这个位置按【真类型】弹真成员，比我们准。
            //
            //   位置判据只有一份：Rules.MemberDotCol（标识符 / `)` / `]` 之后，
            //   以及 With 块的 `.成员`）。
            //   ⚠️ 放在 EnsurePool 之前：点号是高频键，这一处既不需要池子也不需要
            //   COM，先判完再决定要不要去取模块全文。这个判断必须放在最前面
            //   （放后面的话，那条"宽限期到、把候选补回来"的补位路径还会走一遍）。
            //   ⚠️ MEMBER_POPUP=1 可以让点号后也出候选；Ctrl+Space 是用户主动
            //   要列表，两种情形照弹。
            if (!manual && !MemberPopup && Rules.MemberDotCol(line, caret) >= 0)
            {
                Hide();
                return false;
            }

            string text = Backend?.ModuleText();
            EnsurePool(snap.Module, text, snap.Line, wordStart, word);

            // ---- `对象.` 位置（只剩 MEMBER_POPUP=1 和 Ctrl+Space 走得到）----
            // 能定出 `root` 的声明类型就按类型收窄到真成员；定不出来就照旧出
            // 全工程候选（口径：宁可给宽的，也不给"查到了类型但
            // 一个成员都没收进来"的空窗）。
            List<string> memberPool = null;
            bool atMemberPos = false;
            if (Rules.MemberRoot(line, caret, out string root, out int dotCol))
            {
                atMemberPos = true;
                memberPool = MemberPool(root, snap);
            }

            if (word.Length == 0)
            {
                if (memberPool != null || ((manual || MemberPopup) && atMemberPos))
                {
                    // 刚打点号 / 手动唤出：列出全部（0 个字符的 query 匹配一切）
                }
                else if (!manual) { Hide(); return false; }
                else wordStart = caret;      // Ctrl+Space 在空位置：列出全部
            }

            // 刚确认过的词不再提示自己（按 Tab 之后光标处那个词）
            if (!manual && QuietFromAccepted(word, snap)) { Hide(); return false; }

            var ranks = BuildRankMap(snap.Module, ProcOf(snap));

            var hits = new List<Candidate>(64);
            foreach (var name in memberPool ?? _poolNames)
            {
                if (string.IsNullOrEmpty(name)) continue;
                FuzzyHit fh;
                if (!Fuzzy.Match(name, word, out fh)) continue;

                // 单字母只认前缀/精确 —— 否则打一个 `e` 会命中几百个
                if (word.Length == 1 && fh.Kind < 3) continue;

                int rank;
                if (memberPool != null)
                {
                    rank = 3;                     // 成员：天然属于当前对象
                }
                else if (!ranks.TryGetValue(name, out rank))
                {
                    // 声明出来的名字查不到档位 = 当前位置根本看不见它
                    // （别的过程的局部变量 / 别的模块的 Private）—— 直接剔除。
                    if (_poolScoped.Contains(name)) continue;
                    rank = 2;
                }

                var c = new Candidate
                {
                    Name = name,
                    Pos = fh.First,
                    Span = fh.Span,
                    Hits = fh.Positions,
                    HitKind = fh.Kind,
                    ScopeKind = rank,
                    RecentRank = Recent != null ? Recent.Rank(name) : 0,
                };
                c.Kind = Classify(c.Name);
                hits.Add(c);
            }

            if (hits.Count == 0) { Hide(); return false; }

            hits.Sort(Cmp);

            // （对标 IDEA）：命中太多就截断。打一两个字母很容易命中上百个
            //   —— 滚动条长得没有意义，每次重画也白费。排序已经保证"最像的"
            //   在前 240 条里，继续输入一个字母就能收窄。
            //   截断只发生在**给 UI 的列表**上；自检要看全部命中就看 TruncatedFrom。
            const int MaxCandidates = 240;
            TruncatedFrom = hits.Count > MaxCandidates ? hits.Count : 0;
            if (hits.Count > MaxCandidates)
                hits.RemoveRange(MaxCandidates, hits.Count - MaxCandidates);
            // 完全相同的名字只留一条（同名过程 / 同名隐式变量），保留排在前面的那条
            var dedup = new List<Candidate>(hits.Count);
            var seenName = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in hits) if (seenName.Add(h.Name)) dedup.Add(h);
            hits = dedup;

            _wordComplete = word.Length > 0 &&
                            Rules.PrefixIsResolved(word, hits.Select(x => x.Name).ToList());

            if (!manual && HideOnComplete && _wordComplete)
            {
                // 打全了：候选一并不留 —— 否则按 Tab 会把已经打完的名字再补一遍
                Hide();
                return false;
            }

            _items = hits;
            _sel = 0;
            _top = 0;
            // 记下"这批候选属于哪个位置"，Accept 时复核（见 AcceptName）。
            _ctxModule = snap.Module ?? "";
            _ctxLine = snap.Line;
            _ctxWordStart = wordStart;
            Visible = true;
            OnShow?.Invoke(_items, _sel, _top);
            return true;
        }

        /// <summary>光标所在的过程名（ScopeKey）；不在任何过程里返回 null。</summary>
        private string ProcOf(Snapshot snap)
        {
            if (snap == null) return null;
            foreach (var ms in _scopes)
            {
                if (!string.Equals(ms.Module, snap.Module, StringComparison.OrdinalIgnoreCase)) continue;
                var p = Parser.ProcOfLine(ms.Procs, snap.Line);
                return p == null ? null : p.ScopeKey;
            }
            return null;
        }

        /// <summary>
        /// `root.` 的成员名单；**类型不确定时返回 null**（调用方据此整个让给 VBE）。
        ///
        /// 类型来源只有两条，都是代码里写死的：`Dim c As Cls` / `Set c = New Cls`。
        /// 刻意不做推断（返回值类型、With 块、宿主对象）—— 猜错就会弹出
        /// "这个对象根本没有的成员"，比不弹更糟。
        /// </summary>
        private List<string> MemberPool(string root, Snapshot snap)
        {
            if (string.IsNullOrEmpty(root) || snap == null) return null;
            string curModule = snap.Module ?? "";
            string typeName;
            // `Me.` / `this.` 是"本模块自己"，成员含 Private；
            //   别的 root 是"别人的类型"，只能看 Public。
            bool self = string.Equals(root, "Me", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(root, "this", StringComparison.OrdinalIgnoreCase);

            if (self)
            {
                typeName = curModule;
            }
            else
            {
                string proc = ProcOf(snap);
                typeName = null;
                foreach (var ms in _scopes)
                {
                    if (!string.Equals(ms.Module, curModule, StringComparison.OrdinalIgnoreCase)) continue;
                    typeName = ms.TypeOf(root, proc);
                    break;
                }
            }
            if (string.IsNullOrEmpty(typeName)) return null;

            var outp = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1) 同名模块（类模块 / 窗体 / 标准模块）的成员：
            //    Me -> 含 Private；别的模块 -> 只取 Public。
            foreach (var ms in _scopes)
            {
                if (!string.Equals(ms.Module, typeName, StringComparison.OrdinalIgnoreCase)) continue;
                foreach (var kv in (self ? ms.Members() : ms.Exported()))
                    if (seen.Add(kv.Key)) outp.Add(kv.Key);
                return outp.Count > 0 ? outp : null;
            }

            // 2) `Type Point2` 里的成员（挂的是 "type:Point2" 伪过程键）
            foreach (var ms in _scopes)
                foreach (var d in ms.Decls)
                {
                    if (d.Proc == null || !d.Proc.StartsWith("type:", StringComparison.Ordinal)) continue;
                    if (!string.Equals(d.Proc.Substring(5), typeName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (seen.Add(d.Name)) outp.Add(d.Name);
                }

            return outp.Count > 0 ? outp : null;
        }

        /// <summary>"这个词是不是我们刚按 Tab 写进去的"。对不上就当场作废。</summary>
        private bool QuietFromAccepted(string word, Snapshot snap)
        {
            if (_acceptedWord == null) return false;
            if (Environment.TickCount64 - _acceptedTicks > AcceptQuietMs)
            {
                _acceptedWord = null;
                return false;
            }
            bool same = string.Equals(word, _acceptedWord, StringComparison.OrdinalIgnoreCase)
                        && snap.Line == _acceptedLine
                        && string.Equals(snap.Module ?? "", _acceptedModule ?? "",
                                         StringComparison.OrdinalIgnoreCase);
            if (same) return true;
            // 用户动过手（换行 / 换模块 / 改词）=> 记忆作废，下次照常提示
            _acceptedWord = null;
            return false;
        }

        /// <summary>
        /// 排序键（全部"越大越靠前"）：
        ///   命中质量(kind) > 首个命中位置 > 命中跨度
        ///     > 作用域就近 > 最近使用 > 名字长度(短在前)
        /// </summary>
        public static int Cmp(Candidate a, Candidate b)
        {
            int c = b.HitKind.CompareTo(a.HitKind); if (c != 0) return c;
            c = a.Pos.CompareTo(b.Pos); if (c != 0) return c;
            c = a.Span.CompareTo(b.Span); if (c != 0) return c;
            c = b.ScopeKind.CompareTo(a.ScopeKind); if (c != 0) return c;
            c = b.RecentRank.CompareTo(a.RecentRank); if (c != 0) return c;
            c = a.Name.Length.CompareTo(b.Name.Length); if (c != 0) return c;
            return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 图标档位。起池子里只剩自己写的名字，于是档位只有三种：
        /// 过程（func）/ 属性（prop）/ 变量（var）。类型的图标不参与分类。
        /// </summary>
        public string Classify(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            string kind;
            if (_procKinds.TryGetValue(name, out kind))
                return kind == "property" ? "prop" : "func";
            return "var";
        }

        /// <summary>
        /// 按 Tab 确认之后，后面该跟什么：(补括号, 补空格)。
        /// 口径：Sub 补空格 / Function 补括号 / Property 都不补；
        /// 例外：在 Function 体内给自己赋值返回值时不补括号。
        /// 内建函数与关键字已不在候选里，所以只有"自己写的过程"有后续。
        /// </summary>
        public void Followup(string name, string enclosingProc, out bool wantParens, out bool wantSpace)
        {
            wantParens = false;
            wantSpace = false;
            if (string.IsNullOrEmpty(name)) return;

            string kind;
            if (!_procKinds.TryGetValue(name, out kind)) return;

            if (kind == "sub") { wantSpace = true; return; }
            if (kind == "function")
            {
                if (IsOwnProcName(name.ToLowerInvariant(), enclosingProc)) return;   // 给返回值赋值，不补
                wantParens = true;
                return;
            }
            // property：什么都不补
        }

        public static bool IsOwnProcName(string name, string enclosingProc)
        {
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(enclosingProc)) return false;
            string b = enclosingProc.ToLowerInvariant();
            int hash = b.IndexOf('#');
            if (hash >= 0) b = b.Substring(0, hash);
            return name.ToLowerInvariant() == b;
        }

        // ------------------------------------------------------------------
        // 选择
        // ------------------------------------------------------------------

        public void Move(int delta)
        {
            if (!Visible || _items.Count == 0) return;
            _sel = (_sel + delta) % _items.Count;
            if (_sel < 0) _sel += _items.Count;
            ScrollToSelected();
            OnMove?.Invoke();
        }

        public void SetSelected(int index)
        {
            if (_items.Count == 0) return;
            _sel = Math.Max(0, Math.Min(_items.Count - 1, index));
            ScrollToSelected();
            OnMove?.Invoke();
        }

        public int PageSize => Math.Max(1, ViewRows - 1);

        public void ScrollToSelected()
        {
            if (_sel < _top) _top = _sel;
            else if (_sel >= _top + ViewRows) _top = _sel - ViewRows + 1;
            int maxTop = Math.Max(0, _items.Count - ViewRows);
            if (_top > maxTop) _top = maxTop;
            if (_top < 0) _top = 0;
        }

        public void ScrollBy(int lines)
        {
            if (!Visible) return;
            int maxTop = Math.Max(0, _items.Count - ViewRows);
            _top = Math.Max(0, Math.Min(maxTop, _top + lines));
            OnMove?.Invoke();
        }

        /// <summary>滚轮 / 拖滚动条：直接指定顶行。</summary>
        public void SetTop(int top)
        {
            if (!Visible) return;
            int maxTop = Math.Max(0, _items.Count - ViewRows);
            _top = Math.Max(0, Math.Min(maxTop, top));
            if (_sel < _top) _sel = _top;
            else if (_sel >= _top + ViewRows) _sel = _top + ViewRows - 1;
            OnMove?.Invoke();
        }

        public void Hide()
        {
            if (!Visible) return;
            Visible = false;
            _items = new List<Candidate>();
            _sel = 0; _top = 0;
            _wordComplete = false;
            OnHide?.Invoke();
        }

        /// <summary>接受选中的候选：写回 VBE，然后收起窗口。</summary>
        public bool Accept()
        {
            if (!Visible || _items.Count == 0) return false;
            var cand = _items[_sel];
            return AcceptName(cand.Name);
        }

        public bool AcceptName(string name)
        {
            if (string.IsNullOrEmpty(name) || Backend == null) { Hide(); return false; }

            // 写完之前重新取一次快照：候选项算出来之后用户可能又敲了字，
            //   拿旧的 word_start_col 去替换会把行改坏。
            var snap = Backend.Snapshot();
            if (snap == null) { Hide(); return false; }

            int wordStart;
            string acceptLine = snap.Text ?? "";
            int acceptCaret = Rules.CharColFromDisp(acceptLine, snap.CaretCol);
            Rules.WordBefore(acceptLine, acceptCaret, out wordStart);
            int endCol = acceptCaret;
            if (wordStart > endCol) wordStart = endCol;
            // （对标 IDEA）：补全替换的是**整个标识符**，不是"光标左边那一截"。
            //   光标停在 `MsgB|ox` 中间按 Tab 时，只换 [wordStart, caret) 会
            //   结果是 `MsgBoxox`。IDEA 在这里也是整词替换 —— 右侧连续的标识符字符
            //   一并吃掉（遇到 `(` `.` `,` 等就停）。
            while (endCol <= acceptLine.Length && Rules.IsIdentChar(acceptLine[endCol - 1]))
                endCol++;

            // 复核"现在这个位置还是算候选时那个位置吗"。
            //   候选窗开着时按 Home/End、或鼠标点到同一行的别处，行文本没变、
            //   轮询不重触发，列表里还是旧候选；此时按 Tab 会拿**新**的 wordStart
            //   去替换 —— 按过 Home 就变成"把候选插到行首"，点到词中间就变成
            //   "把那个词截成两半"。位置对不上就收起窗口，把 Tab 还给 VBE。
            //   注意只比"行号 / 取词起点 / 模块"，不比词本身：用户多打一个字母
            //   是正常的（`MsgB` → Tab），那时 wordStart 不变、必须放行。
            if (_ctxModule != null &&
                (snap.Line != _ctxLine ||
                 wordStart != _ctxWordStart ||
                 !string.Equals(snap.Module ?? "", _ctxModule, StringComparison.OrdinalIgnoreCase)))
            {
                Hide();
                return false;
            }

            string enclosing = null;
            foreach (var ms in _scopes)
            {
                if (!string.Equals(ms.Module, snap.Module, StringComparison.OrdinalIgnoreCase)) continue;
                var p = Parser.ProcOfLine(ms.Procs, snap.Line);
                enclosing = p == null ? null : p.Name;
                break;
            }
            if (enclosing == null) enclosing = Backend.EnclosingProcName();

            Followup(name, enclosing, out bool wantParens, out bool wantSpace);
            bool ok = Backend.ApplyCompletion(snap.Line, wordStart, endCol, name,
                wantParens, wantSpace, enclosing);
            Hide();
            if (ok)
            {
                // 记一笔"刚写进去的是这个词" —— 轮询下一拍发现文档变了会再触发一次，
                // 那时光标处正是这个词，不挡住的话用户会看到自己刚确认的名字又被提示。
                _acceptedWord = name;
                _acceptedModule = snap.Module ?? "";
                _acceptedLine = snap.Line;
                _acceptedTicks = Environment.TickCount64;
                // 这里只记一笔，**不**立刻写盘 —— 见 App.PollOnce 里的合并 Flush。
                //   原来每接受一次候选就 WriteAllText + File.Move 一次，家目录在
                //   OneDrive / 网络盘时确认完会明显卡一下，而且记录只增不减。
                Recent?.Note(name);
                InvalidatePool();
            }
            return ok;
        }

    }
}
