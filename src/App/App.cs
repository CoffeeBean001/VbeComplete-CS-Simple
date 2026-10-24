// 主流程：全局钩子 + 自适应轮询 + 按键路由 + 各种"改一行"的操作。
//
// 这是整个工具唯一允许"决定吃不吃一个按键"的地方，所以三条铁律在这里落地：
//   1. 钩子线程只读内存里的状态（_codePaneFocused / _popupVisible / _lastKeyTs），
//      一次 COM 都不许碰 —— 碰一次就可能让全系统的键盘卡住；
//   2. 真动作一律投递到 UI 线程执行；
//   3. 吃掉的键如果最后没做成事，必须把键还回去（绝不"吞了但什么都没发生"）。
//
// CPU 上花的功夫也在这里：
//   * 前台不是 VBE 代码窗时【完全不调 COM】—— 用户在 Excel 表格里、
//     在浏览器里的时候，我们一个 COM 都不发；
//   * 在代码窗里也按"刚敲过键没有"来调频率：打字时 50ms，停手 200ms。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    internal sealed class App : ApplicationContext
    {
        private const string CodePaneClass = "VbaWindow";
        private const string VbeFrameClass = "wndclass_desked_gsk";
        private static readonly string[] VbePopupClasses = { "NameListWndClass", "PopupTipWndClass" };

        private readonly VbeBackend _be = new VbeBackend();
        private readonly Recent _recent = new Recent();
        private long _recentFlushAt;                      // 上次把 recent.json 写盘的时刻
        private readonly Engine _eng;
        private readonly PopupForm _popup = new PopupForm();
        private readonly Control _disp = new Control();

        private System.Windows.Forms.Timer _timer;
        private NotifyIcon _tray;
        private EventWaitHandle _showEvent;
        private Thread _showThread;

        // ---- 跨线程共享的状态（钩子线程只读，UI 线程写） ----
        private volatile bool _codePaneFocused;
        private volatile bool _popupVisible;
        private long _lastYieldLog;               // "让位"日志限速（每 2 秒最多一条）
        private long _lastJumpLog;                // "跳行不弹"日志限速（每 2 秒最多一条）
        private volatile bool _running = true;
        private volatile Snapshot _hookSnap;      // 最近一次轮询拿到的快照（钩子线程只读）
        private long _popupBeat;                  // 候选窗心跳（UI 线程写，钩子线程只读）
        private volatile int _pendingKeys;        // 按了"可能写字"的键、文档却没变（输入法组字）
        private long _pendingAt;                  // 上面那次计数的时刻（超时就当 0）
        private long _hookSnapTicks;
        // 片段窗的 HWND。必须由 UI 线程在 Show 之前存好、钩子线程只读这个
        // IntPtr：钩子线程去碰 `_snippetForm.Handle` 会在句柄还没建出来时
        // 触发一次【在钩子线程上创建窗口句柄】，那是跨线程建窗，必炸。
        private volatile IntPtr _snippetHwnd;
        private long _lastKeyTs;

        // ---- 探针专用（--enterprobe）：默认 0 = 注入键一律不理，与生产行为一致 ----
        // 为什么需要它：真实的按键路径是"钩子 → 判据 → Post → COM"，而 SendInput
        //   发出来的键带 LLKHF_INJECTED 标记，钩子会（也必须）跳过它们 —— 否则我们
        //   自己补的空格 / 括号会自己触发自己。真机探针要验的正是这条路，于是给它
        //   开一个**只对测试键**生效的放行口：0 = 全不理，-1 = 全放行（模拟连续打字），
        //   其它值 = 只放行这一个 vk。
        internal static volatile int AcceptInjectedVk;

        // ---- 探针用的只读视图（生产路径一个都不读） ----
        internal bool CodePaneFocused => _codePaneFocused;
        internal bool PopupVisible => _popupVisible;
        internal int PendingKeys => _pendingKeys;
        internal Snapshot HookSnap => _hookSnap;
        internal long HookSnapAgeMs => Environment.TickCount64 - _hookSnapTicks;
        internal void RunOnUi(Action a) { try { _disp.Invoke(a); } catch { } }
        internal void RequestExit() { try { _disp.BeginInvoke((Action)ExitThread); } catch { } }

        // ---- 钩子 ----
        private Thread _hookThread;
        private IntPtr _hookHandle = IntPtr.Zero;
        private uint _hookThreadId;
        private Native.LowLevelKeyboardProc _hookProc;   // 必须保持引用，否则被 GC
        // 已吞掉 keydown 的键 -> 吞键时刻。必须是"带时标"的字典而不是 HashSet：
        // 吞了 keydown 之后如果那个 keyup 没回来（切窗口、被别的钩子劫走、程序被挂起），
        // 键就永远留在表里，下一次同键的 keyup 会在**任何**窗口里被吞掉 —— 丢一个
        // keyup 的后果是修饰键粘住（Shift 按一下就一直是选定状态）。所以只认
        // 距今 1.5 秒内的配对，超时当没这回事。
        private readonly Dictionary<int, long> _swallowed = new Dictionary<int, long>();
        private const long SwallowPairMs = 1500;

        // ---- 轮询状态 ----
        private Snapshot _lastSnap;
        private string _lastModName = "";
        private int _shutdownDone;                // Shutdown 的幂等标志（Interlocked）

        // ---- 代码片段 ----
        private SnippetStore _snippetStore;
        private SnippetForm _snippetForm;

        private readonly bool NoCtrlNav = Settings.Flag("NO_CTRL_NAV", false);
        private readonly bool NoSnippets = Settings.Flag("NO_SNIPPETS", false);
        private readonly bool NoComment = Settings.Flag("NO_COMMENT", false);
        private readonly bool NoDupLine = Settings.Flag("NO_DUP_LINE", false);
        private readonly bool NoDelLine = Settings.Flag("NO_DEL_LINE", false);
        private readonly bool NoInsertAbove = Settings.Flag("NO_INSERT_ABOVE", false);
        private readonly bool NoEnterIndent = Settings.Flag("NO_ENTER_INDENT", false);
        private readonly bool NoAutoClose = Settings.Flag("NO_AUTOCLOSE", false);
        private readonly bool NoBranchAlign = Settings.Flag("NO_BRANCH_ALIGN", false);
        private readonly bool NoShiftNav = Settings.Flag("NO_SHIFT_NAV", false);
        private readonly bool NoAutoPair = Settings.Flag("NO_AUTOPAIR", false);
        private readonly bool NoBsDelLine = Settings.Flag("NO_BS_DELLINE", false);
        private readonly bool NoBsDelPair = Settings.Flag("NO_BS_DELPAIR", false);
        // 自动弹窗只认按键：鼠标右键粘贴 / 拖拽 / 别的程序改代码造成的文本变化
        // 只收起、不弹（REQUIRE_KEY_TRIGGER 打开时，默认开）。
        private readonly bool RequireKeyTrigger = Settings.Flag("REQUIRE_KEY_TRIGGER", true);
        private const long KeyTriggerWindowMs = 1200;

        // 最近一次剪贴板按键（Ctrl+C/X/V、Shift+Insert）的时刻。粘贴落地的那一次
        // 文档变化必须按"内容变了"处理、不能按"打字"处理：复制一个变量名粘到别处，
        // 光标正好停在一个像标识符的结尾上，按打字算就会弹出一串候选。
        private long _clipAt;
        private const long ClipGuardMs = 1000;

        public App()
        {
            // 强制建句柄：跨线程 BeginInvoke 需要一个真窗口
            var _ = _disp.Handle;

            _be.SendChar = ch => Native.SendUnicodeChar(ch);
            _be.SendVk = (vk, ctrl, shift) => Native.SendVk(vk, ctrl, shift);

            // 不再载入词库。候选池只由「工程里的声明」构成（见 Engine.EnsurePool），
            //   启动时零磁盘 IO、零 SQLite 依赖。
            _eng = new Engine
            {
                Backend = _be,
                Recent = _recent,
            };
            _eng.OnShow = (items, sel, top) =>
            {
                _popup.SetData(items, sel, top, _eng.ViewRows);
                ShowPopupAtCaret();
                // 镜像标志必须跟着【真实可见性】走：ShowPopupAtCaret 拿不到插入符
                //   时会 _eng.Hide()，此时窗口压根没显示。无条件置 true 会让
                //   _popupVisible 永久为真 —— 之后 ↑/↓/PgUp/Esc 全被吞掉却什么都
                //   不发生（且 Engine.Hide 的 `if (!Visible) return` 让 OnHide
                //   再也纠正不了它）。
                _popupVisible = _popup.Visible;
                _popupBeat = Environment.TickCount64;
                if (!_popupVisible) _eng.Hide();
            };
            _eng.OnHide = () =>
            {
                _popupVisible = false;
                _popup.ResetPlacement();
                if (_popup.Visible) _popup.Hide();
            };
            _eng.OnMove = () =>
            {
                _popup.UpdateView(_eng.SelectedIndex, _eng.TopIndex);
            };

            _popup.RowClicked += i => { _eng.SetSelected(i); _eng.Accept(); };
            _popup.Wheel += lines => _eng.ScrollBy(lines);
            _popup.ScrollTo += top =>
            {
                _eng.SetTop(top);
                _popup.UpdateView(_eng.SelectedIndex, _eng.TopIndex);
            };
            // VBE 自己的成员列表弹出来的时候我们让位（它的类型信息比我们准）
            _eng.YieldToVbe = () => Native.VbeMemberListVisible();
        }

        public void Start()
        {
            // 开机自启：按 autostart.txt 里记的意图把注册表修正一遍（第一次运行默认开）。
            // 必须放在 SetupTray 之前：托盘菜单里的勾选状态是在建菜单时读注册表的，
            //   顺序反过来会看到"已经登记了但菜单上没勾"。
            try { AutoStart.Sync(); } catch (Exception e) { Log.Write("开机自启同步失败：{0}", e.Message); }
            SetupTray();
            StartHook();
            StartShowListener();
            _timer = new System.Windows.Forms.Timer { Interval = 200 };
            _timer.Tick += (s, e) => PollOnce();
            _timer.Start();
            Log.Write("已启动（Home={0}，{1}）", Settings.Home, AutoStart.Describe());

            // 托盘图标刚注册时，Windows 才刚建好它在 NotifyIconSettings 里的那条记录，
            // 所以延后 2.5 秒再去找 —— 太早会找不到（第一次运行尤其如此）。
            var pin = new System.Windows.Forms.Timer { Interval = 2500 };
            pin.Tick += (s, e) =>
            {
                pin.Stop();
                pin.Dispose();
                try { TrayPin.EnsureVisible(); } catch { }
            };
            pin.Start();

            // 启动了就必须让用户看得见。这个程序没有主窗口，图标又默认折叠进
            // Win11 的托盘溢出区 —— 不给提示的话，"双击了没反应"和"已经正常运行"
            // 在屏幕上长得一模一样（用户就是这么报的 bug）。
            // 只在第一次运行时提示：现在它会开机自启，每次登录都弹一个气泡就是骚扰。
            //   之后手动再启动时，若已有一份在跑，SignalExisting() 会负责提示。
            string first = Path.Combine(Settings.Home, "firstrun.txt");
            if (!File.Exists(first))
            {
                ShowBalloon("VbeComplete 已启动",
                    "没有主窗口：图标在右下角托盘里（若看不到，点一下任务栏的 ^ ，把它拖到外面即可常驻）。\r\n"
                    + "右键那个图标可以开关「开机自动启动」或退出。");
                try { File.WriteAllText(first, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")); } catch { }
            }
        }

        private void ShowBalloon(string title, string text)
        {
            try
            {
                if (_tray == null) return;
                _tray.BalloonTipTitle = title;
                _tray.BalloonTipText = text;
                _tray.BalloonTipIcon = ToolTipIcon.Info;
                _tray.ShowBalloonTip(5000);
            }
            catch (Exception e) { Log.Write("气泡提示失败：{0}", e.Message); }
        }

        /// <summary>
        /// 监听"又有人启动了第二个实例"。Program 那边发现互斥体已存在时会 Set
        /// 这个事件，我们收到就把托盘图标的存在感刷一遍 —— 否则用户双击后的
        /// 全部观感就是"什么都没发生"。
        /// </summary>
        private void StartShowListener()
        {
            try
            {
                _showEvent = new EventWaitHandle(false, EventResetMode.AutoReset, Program.ShowEventName);
                _showThread = new Thread(() =>
                {
                    while (_running)
                    {
                        try
                        {
                            if (!_showEvent.WaitOne(500)) continue;
                            if (!_running) return;
                            // 必须回 UI 线程碰 NotifyIcon
                            if (_disp.InvokeRequired)
                                _disp.BeginInvoke((Action)(() => ShowBalloon(
                                    "VbeComplete 已经在运行",
                                    "不用再启动一次。图标在右下角托盘的溢出区里，右键可以退出。")));
                            else
                                ShowBalloon("VbeComplete 已经在运行",
                                    "不用再启动一次。图标在右下角托盘的溢出区里，右键可以退出。");
                        }
                        catch { return; }
                    }
                })
                { IsBackground = true, Name = "show-listener" };
                _showThread.Start();
            }
            catch (Exception e) { Log.Write("重复启动监听未装上：{0}", e.Message); }
        }

        /// <summary>
        /// 托盘图标：没有它，用户只能去任务管理器杀进程。
        /// 图标必须用【系统自带】的 —— 不引入任何资源文件，保持单文件干净。
        /// </summary>
        private void SetupTray()
        {
            try
            {
                var menu = new ContextMenuStrip();
                var title = new ToolStripMenuItem("VbeComplete · VBA 补全") { Enabled = false };
                menu.Items.Add(title);
                menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add("代码片段…", null, (s, e) =>
                {
                    try { ToggleSnippets(); }
                    catch (Exception ex) { Log.Write("托盘打开片段窗失败：{0}", ex.Message); }
                });

                // 开机自启：勾选状态按注册表的**真实情况**显示（不是按配置），
                // 免得用户看到"勾着但其实早被清理软件删了"。
                var miAutoStart = new ToolStripMenuItem("开机自动启动")
                {
                    CheckOnClick = false,                    // 自己处理，避免和真实状态脱节
                    Checked = AutoStart.IsRegistered(out _),
                    ToolTipText = "随 Windows 登录自动运行（写 HKCU 的 Run 键，不需要管理员）",
                };
                miAutoStart.Click += (s, e) =>
                {
                    bool want = !AutoStart.IsRegistered(out _);
                    bool ok = AutoStart.SetEnabled(want);
                    miAutoStart.Checked = AutoStart.IsRegistered(out _);
                    Log.Write("托盘切换开机自启 -> {0}：{1}", want ? "开" : "关", ok ? "成功" : "失败");
                    if (!ok)
                        ShowBalloon("开机自启没设置成功",
                            "写注册表被拒绝了（常见于安全软件拦截）。日志里有详细原因。");
                };
                menu.Items.Add(miAutoStart);
                menu.Items.Add("打开日志", null, (s, e) =>
                {
                    // 路径必须加引号 —— 家目录带空格（`D:\我的 工具\…`）时
                    //   记事本会把路径切成两段，这条菜单恰恰是用户排查问题的唯一入口。
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo(
                            "notepad.exe", "\"" + Log.Path_ + "\"") { UseShellExecute = true };
                        System.Diagnostics.Process.Start(psi);
                    }
                    catch (Exception ex) { Log.Write("打开日志失败：{0}", ex.Message); }
                });
                menu.Items.Add("退出", null, (s, e) => Shutdown());

                _tray = new NotifyIcon
                {
                    // 一定要用自己的图标：SystemIcons.Application 是一张通用的
                    //   "空白窗口"，在托盘里跟十几个程序长得一模一样，用户根本找不到。
                    Icon = IconAssets.Tray(),
                    Text = "VbeComplete · VBA 补全",
                    Visible = true,
                    ContextMenuStrip = menu,
                };
                _tray.DoubleClick += (s, e) =>
                {
                    var snap = _be.Snapshot();
                    Log.Write("托盘双击：snapshot={0}", snap == null ? "null" : snap.ToString());
                };
            }
            catch (Exception e) { Log.Write("托盘创建失败：{0}", e.Message); }
        }

        // ------------------------------------------------------------------
        // 钩子线程
        // ------------------------------------------------------------------

        private void StartHook()
        {
            _hookThread = new Thread(HookLoop)
            {
                IsBackground = true,
                Name = "kbd-hook",
            };
            _hookThread.Start();
        }

        private void HookLoop()
        {
            _hookThreadId = Native.GetCurrentThreadId();
            _hookProc = HookProc;
            IntPtr mod = Native.GetModuleHandle(null);
            _hookHandle = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, _hookProc, mod, 0);
            if (_hookHandle == IntPtr.Zero)
            {
                Log.Write("装钩子失败，错误码 {0}", Marshal.GetLastWin32Error());
                return;
            }
            Log.Write("钩子已装上（tid={0}）", _hookThreadId);

            Native.MSG msg;
            while (Native.GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
            {
                Native.TranslateMessage(ref msg);
                Native.DispatchMessage(ref msg);
            }
            Native.UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
        }

        /// <summary>
        /// 钩子回调的**唯一**入口：只负责把异常挡在 native 边界之内。
        /// <para>整个回调必须兜住异常。托管异常一旦穿过
        /// <c>CallNextHookEx</c> 的调用帧，CLR 会直接 FailFast 终止进程 ——
        /// 用户看到的是"托盘图标突然没了、日志里一个字都没有"，而且已经被
        /// Swallow 的按键也不会还原。这里兜住：记一笔、把键原样放行。</para>
        /// </summary>
        private IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try { return HookProcCore(nCode, wParam, lParam); }
            catch (Exception e)
            {
                try { Log.Write("钩子回调异常（已拦下，按键放行）：{0}", e); } catch { }
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }
        }

        private IntPtr HookProcCore(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0) return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            int msg = wParam.ToInt32();
            bool isDown = msg == Native.WM_KEYDOWN || msg == Native.WM_SYSKEYDOWN;
            bool isUp = msg == Native.WM_KEYUP || msg == Native.WM_SYSKEYUP;
            if (!isDown && !isUp) return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            var kb = (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(
                lParam, typeof(Native.KBDLLHOOKSTRUCT));

            // 注入的键（我们自己发的空格/括号/回车）不算"用户按键"：否则会自己触发自己
            // （探针 --enterprobe 通过 AcceptInjectedVk 单独放行它要发的那几个键）
            if ((kb.flags & Native.LLKHF_INJECTED) != 0 &&
                AcceptInjectedVk != -1 && (int)kb.vkCode != AcceptInjectedVk)
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            int vk = (int)kb.vkCode;

            // ---- keyup：跟被吞掉的 keydown 配对的也一起吞 ----
            // 只吞 keydown 的话，VBE 会收到一个没有 down 的 up
            //   （Ctrl+D / Ctrl+K 这类会被当成"半截按键"），表现为偶发的重复动作。
            if (isUp)
            {
                lock (_swallowed)
                {
                    // 只有"刚吞了 keydown 的那个键"才连带吞掉它的 keyup；过期的条目
                    // 一并清掉，绝不让它去吞不相干的 keyup（丢 keyup = 修饰键粘住）。
                    if (_swallowed.TryGetValue(vk, out long at))
                    {
                        _swallowed.Remove(vk);
                        if (Environment.TickCount64 - at <= SwallowPairMs) return (IntPtr)1;
                    }
                }
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }

            // 时标打在【这里】：放到别的回调里就晚了（那时钩子已经把事件
            //   post 给监听器，文本可能已经变了）
            _lastKeyTs = Environment.TickCount64;

            bool ctrl = Native.Down(Native.VK_CONTROL);
            bool shift = Native.Down(Native.VK_SHIFT);
            bool alt = Native.Down(Native.VK_MENU);
            bool win = Native.Down(Native.VK_LWIN) || Native.Down(Native.VK_RWIN);

            // 不在 VBE 代码窗（或片段窗正占着焦点）：一个键都不碰。
            // 片段窗那道闸必须看【实时前台】而不是"窗子开着就算"：
            //   `_codePaneFocused` 是轮询算的，片段窗刚弹出来的那一两百毫秒里它还
            //   停在 true —— 这时候在搜索框里打字会被当成在代码里打字（敲个 `(`
            //   就往模块里补一对括号、敲回车就跑"自动缩进"）。反过来，只看
            //   "窗子开着"又会把"用户点了下代码窗继续写、顺手 Ctrl+Q 收起"一起
            //   冻住。GetForegroundWindow 是纯 Win32（微秒级），每次按键问一次没成本。
            // 但如果按的是我们"该管"的组合键却因为焦点判定被跳过，就记一笔 ——
            //   这正是"快捷键全都不灵了"最常见的真因（焦点探测说不在代码窗，
            //   于是连一次路由都没走）。只在 LOG=1 时落盘，平时零成本。
            bool snipActive = SnippetFormActive();
            if (!_codePaneFocused || snipActive)
            {
                if (ctrl && IsHotkeyDiagVk(vk))
                    Log.Write("热键被跳过：焦点判定=不在 VBE 代码窗（vk=0x{0:X2} shift={1} codePane={2} snippet={3}）",
                        vk, shift, _codePaneFocused, snipActive);
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }

            if (win) return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            // ---- 剪贴板：复制 / 剪切 / 粘贴一律不算"打字" ----
            // 粘贴是"原样放进来"，不是"用户敲完了一个词"。复制一个变量名再粘贴，
            //   光标恰好停在一个像标识符的结尾上，轮询会看到"刚按过键 + 文本变了"
            //   而弹出一串候选。这里只打一个标记，让轮询把随之而来的那一次变化
            //   当成外部改动：收窗、重建池子，但不弹。
            // 按键本身照常放行 —— 粘贴由 VBE 自己做，我们不碰。
            if (Native.IsClipboardChord(ctrl, shift, alt, vk))
            {
                _clipAt = Environment.TickCount64;
                if (_popupVisible) Post(() => _eng.Hide());
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }

            // ---- 数"按了可能写字的键、文档却没变"的次数 ----
            // 那就是输入法正在组字的信号（拼音被吃掉、还没上屏）。组字状态下
            // 回车是【上屏键】、退格是【删拼音】，抢了就是"字没上屏反而多一行"。
            // 文档一变（轮询看到）计数就清零，所以正常打字时它总是 0。
            if (Editing.IsTextKey(vk) && !ctrl && !alt)
            {
                _clipAt = 0;        // 真的开始打字了，粘贴那次宽限期就此作废
                int n = _pendingKeys;
                if (n < 99) _pendingKeys = n + 1;
                _pendingAt = Environment.TickCount64;
            }

            // ==============================================================
            // 1. 候选窗可见时的导航键
            // ==============================================================
            if (PopupAlive && !ctrl && !alt)
            {
                if (vk == Native.VK_DOWN && !shift)
                {
                    Post(() => _eng.Move(1));
                    return Swallow(vk);
                }
                if (vk == Native.VK_UP && !shift)
                {
                    Post(() => _eng.Move(-1));
                    return Swallow(vk);
                }
                // Shift+↑/↓ = 跳出候选列表 + 光标上/下移一行。
                //   刻意【不】走上面"在列表里选词"那条：VBE 原生的 Shift+↑/↓ 是
                //   扩选，不是移光标，所以只能由我们用 COM 落光标（并吞掉按键）。
                if ((vk == Native.VK_UP || vk == Native.VK_DOWN) && shift && !NoShiftNav)
                {
                    int delta = vk == Native.VK_UP ? -1 : 1;
                    Post(() =>
                    {
                        _eng.Hide();
                        if (!_be.MoveCaretLine(delta)) Native.SendVk(vk, false, true);
                    });
                    return Swallow(vk);
                }
                if (vk == Native.VK_NEXT)
                {
                    Post(() => _eng.Move(_eng.PageSize));
                    return Swallow(vk);
                }
                if (vk == Native.VK_PRIOR)
                {
                    Post(() => _eng.Move(-_eng.PageSize));
                    return Swallow(vk);
                }
                // **空格 = "这个词打完了"**，VBE 自己的列表也是在这一下收的。
                //   空格必须当场处理，不能全指望轮询的 changed —— changed 比
                //   的是 **Rstrip 之后**的文本（行尾空白刻意不算文档变化，免得光标
                //   在行尾摆动时反复触发），于是 `Set ` 与 `Set` 在轮询眼里一字不差，
                //   永远走不到 Trigger，`else if (_popupVisible)` 那条只是把窗子挪到
                //   新的光标位置 ⇒ 列表就那么一直挂着。
                //   "set + 空格不收、let + 空格会收" 正是这条：let 是打到
                //   第三个字母时 PrefixIsResolved 判成"打全了"提前收的，跟空格无关；
                //   set 因为工程里还有更长的 `Set…` 名字（或 Reset/LSet 之类），
                //   判不成打全 ⇒ 列表还在，于是空格这一下必须由我们自己收。
                //   只收窗、**不吞键** —— 空格照常落进文档。
                if (vk == Native.VK_SPACE)
                {
                    Post(() => _eng.Hide());
                    return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }
                if (vk == Native.VK_TAB && !shift)
                {
                    Post(() => { if (!_eng.Accept()) Native.SendVk(Native.VK_TAB); });
                    return Swallow(vk);
                }
                if (vk == Native.VK_RETURN && !shift)
                {
                    // 回车【不】确认候选 —— 回车在 VBE 里是换行，误触发代价太大
                    //   （确认键只用 Tab，回车不算）。
                    //   这里只收起列表。
                    // **不要在这里 return**。用户敲完 `If x > 0 Then` 立刻回车时，
                    //   轮询还没看到最后那个 `n`，候选窗正巧还挂着（"The" 匹配到 "Then"），
                    //   在这里 return 就等于把回车整个交回 VBE —— 于是"只换行、不补 End If"。
                    //   这里只收窗、继续往下走，让下面的回车分支自己判。
                    Post(() => _eng.Hide());
                }
                if (vk == Native.VK_ESCAPE)
                {
                    Post(() => _eng.Hide());
                    return Swallow(vk);
                }
                if (vk == Native.VK_LEFT || vk == Native.VK_RIGHT)
                {
                    // 移动光标即放弃本次补全。刻意【不】吞键 —— 光标要照常移动。
                    Post(() => _eng.Hide());
                    return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }
            }

            // ==============================================================
            // 2. 手动唤出
            // ==============================================================
            if (ctrl && !shift && !alt && vk == Native.VK_SPACE)
            {
                Post(() =>
                {
                    var snap = _be.Snapshot();
                    if (snap != null) _eng.Trigger(true, snap);
                });
                return Swallow(vk);
            }

            // ==============================================================
            // 3. Ctrl+Q：代码片段窗（开关）
            //    排在所有"改写文本"的分支之前 —— 它是纯粹的开/关界面动作，
            //    与候选窗是否可见无关。Ctrl+Q 在 VBE 里本来什么也不做。
            // ==============================================================
            if (ctrl && !shift && !alt && vk == 'Q' && !NoSnippets)
            {
                Post(ToggleSnippets);
                return Swallow(vk);
            }

            // ==============================================================
            // 4. 会改写文本的组合键
            // 每一个都必须"做成了才算吃掉了键"。命令返回 false（没快照 / COM
            //   失败 / 没有可落的位置）就把键原样送回去 —— 绝不"吞了但什么都没
            //   发生"—— 那正是"快捷键都没了"时屏幕上的样子。
            // ==============================================================
            // `/` 的识别要放宽一点。"Ctrl+/ 失效"时，最省事的解释
            //   就是 vk 没对上：主键盘 `/` 是 VK_OEM_2，但**小键盘除号**是
            //   VK_DIVIDE；而法语 AZERTY 这类布局里 `/` 本身就带 Shift。
            //   被 `!shift` 那条挡住的话，钩子就当没看见这个键 —— 快捷键没坏，
            //   是我们不认。这里两个 vk 都认、允许 Shift、排除 Alt。
            //   Ctrl+Shift+/ 在 VBE 里没有任何原生用途，多认这一种不会撞车。
            if (ctrl && !alt && Native.IsSlashKey(vk) && !NoComment)
            {
                Post(() =>
                {
                    // 别在失败时补发一个裸 `/`：SendVk 遇到"Ctrl 正被物理按住"
                    //   会刻意不发 Ctrl，于是补发的其实是"/"这个字符，直接往用户
                    //   代码里插一个斜杠 —— 比"什么都没发生"更糟。VBE 里 Ctrl+/
                    //   本来也不做任何事，所以这里只是记一笔。
                    if (!Commands.ToggleComment(_be))
                        Log.Write("Ctrl+/ 没做成：拿不到选区 / 整段全是空行 / 一行都没写成功（键已吞掉，VBE 里本也无此功能）");
                });
                return Swallow(vk);
            }
            // ⚠️ Shift+D / Shift+K 绝不能接管 —— 那是输入大写字母
            if (ctrl && !shift && !alt && vk == 'D' && !NoDupLine)
                return SwallowAndRun('D', true, false, () => Commands.DuplicateLine(_be),
                    replayOnFail: false);
            if (ctrl && !shift && !alt && vk == 'K' && !NoDelLine)
                return SwallowAndRun('K', true, false, () => Commands.DeleteLine(_be),
                    replayOnFail: false);
            if (ctrl && shift && !alt && vk == Native.VK_RETURN && !NoInsertAbove)
                return SwallowAndRun(Native.VK_RETURN, true, true, () => Commands.InsertLineAbove(_be),
                    replayOnFail: false);
            // ==============================================================
            // 4c. Ctrl+↑ / Ctrl+↓（**不带 Shift**）：过程内的两个端点
            //     ↑ 落到本过程【最后一个 Dim 行】的行尾；
            //     ↓ 落到【过程收尾行上方最后一条代码】的行尾。
            //     各对应本过程的一个端点。
            //     带 Shift 的组合原样落到系统/VBE。
            // ==============================================================
            if (ctrl && !shift && !alt && (vk == Native.VK_UP || vk == Native.VK_DOWN) && !NoCtrlNav)
            {
                int dir = vk == Native.VK_UP ? -1 : 1;
                // 判据不成立（光标在声明区 / 过程里没有 Dim / 过程体为空）就把键
                // 还回去。SendVk 在"Ctrl 正被物理按住"时不合成 Ctrl，所以实际
                // 发出的是裸 ↑/↓。
                return SwallowAndRun(vk, true, false, () => Commands.CtrlProcJump(_be, dir));
            }

            // ==============================================================
            // 5. Shift+Enter：在下方新起一行（不拆行），缩进对齐 + 补收尾
            //    ⚠️ 必须排在"回车自动缩进"之前，两条判据互斥（这条要求 Shift）
            // ==============================================================
            if (shift && !ctrl && !alt && vk == Native.VK_RETURN)
                return SwallowAndRun(Native.VK_RETURN, false, true,
                    () => _be.NewLineBelow(false, !NoAutoClose, !NoBranchAlign));

            // ==============================================================
            // 6. 回车自动缩进（缩进 / 补收尾 / 分支对齐）
            //    只有"整行注释 / 块结构开头 + 光标在行尾 + 行内无 Tab"才接管，
            //    其余一律放行给 VBE 原生回车（零风险）。
            // ==============================================================
            if (!ctrl && !shift && !alt && vk == Native.VK_RETURN && !NoEnterIndent)
            {
                string lineText;
                int caretDisp;
                bool fresh = FreshHookSnap(out lineText, out caretDisp);
                if (fresh && Editing.EnterIndentWanted(lineText, caretDisp))
                    return SwallowAndRun(Native.VK_RETURN, false, false,
                        () => _be.NewLineBelow(true, !NoAutoClose, !NoBranchAlign));

                // 快照"落后一两格"时也要接管 —— 否则回车补结构会漏。
                //   用户敲完 `For i = 1 To 3` 的最后一个字符【立刻】回车：那一瞬间
                //   轮询还没把这一格读回来（`_pendingKeys > 0`、行文本少一个 `3`），
                //   上面的判据必然 false ⇒ 回车原样交给 VBE ⇒ 只换行、不补 `Next`。
                //   打字越快越常撞上，所以用户的感觉就是"根本不会补"。
                //
                //   这一档不猜结论，只做"先接管"，然后交给 UI 线程用【真 COM 数据】判：
                //   NewLineBelow 内部会重新读行、重新判（不是块头 / 光标不在行尾 /
                //   后面还有内容 -> 一律返回 false），判不过就把回车原样还回去 ——
                //   绝不"吞了却什么都没发生"。
                //   ⚠️ 输入法正在组字时**绝不接管**：那一下回车是"上屏"，抢了就成了
                //      "字没上屏反而多一行"。这一问走 Win32（ImmGetContext），不碰 COM。
                if (!Native.ImeComposing(Native.FocusedVbeWindow()) &&
                    Editing.LooksLikeBlockHeaderPrefix(lineText) &&
                    caretDisp + 16 >= Rules.DispCol(lineText, (lineText ?? "").Length))
                    return SwallowAndRun(Native.VK_RETURN, false, false,
                        () => _be.NewLineBelow(true, !NoAutoClose, !NoBranchAlign));
            }

            // ==============================================================
            // 7. 退格：删整行空白 / 删一对空括号
            // ==============================================================
            if (!ctrl && !shift && !alt && vk == Native.VK_BACK &&
                (!NoBsDelLine || !NoBsDelPair))
            {
                string lineText;
                int caretDisp;
                if (FreshHookSnap(out lineText, out caretDisp))
                {
                    bool blankLine = !NoBsDelLine && Editing.BlankLineDeleteWanted(lineText, caretDisp);
                    int left, right;
                    bool pair = !NoBsDelPair && !blankLine &&
                                Editing.PairDeleteSpan(lineText, caretDisp, out left, out right);
                    if (blankLine)
                        return SwallowAndRun(Native.VK_BACK, false, false, () => _be.DeleteBlankLineHere());
                    if (pair)
                        return SwallowAndRun(Native.VK_BACK, false, false, () => _be.DeletePairHere());
                }
            }

            // ==============================================================
            // 8. 自动配对括号 / 引号（含中文标点的全角）
            // ==============================================================
            if (NoAutoPair == false && !ctrl && !alt)
            {
                char half = Native.PairCharForKey(vk, shift);
                if (half != '\0')
                {
                    // 到这里才问输入法状态（要发一次窗口消息）。放在"已经确定要
                    //   接管这个键"之后，别每次按 Shift 都去问。
                    char ch = Native.PairCharForIme(half, Native.GetForegroundWindow());
                    Post(() =>
                    {
                        _eng.Hide();
                        if (!_be.InsertPair(ch)) Native.SendUnicodeChar(ch);
                    });
                    return Swallow(vk);
                }
            }

            // 只记录"按了 Ctrl 却没命中任何分支"的、与该管的热键相关的键 ——
            // 用来排查"快捷键不生效"。日常打字（Ctrl+C / Ctrl+V 之类）不进日志。
            if (ctrl && IsHotkeyDiagVk(vk))
                Log.Write("热键未命中：vk=0x{0:X2} ctrl={1} shift={2} alt={3} popup={4}",
                    vk, ctrl, shift, alt, _popupVisible);

            return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        /// <summary>吃掉这一下按键（keydown），并记下 vk 好把配对的 keyup 也吞掉。</summary>
        private IntPtr Swallow(int vk)
        {
            long now = Environment.TickCount64;
            lock (_swallowed)
            {
                // 顺手清理过期条目：正常情况下 keyup 一来就删掉了，这里只是兜底，
                // 免得"吞了 keydown 却没有 keyup"的残渣越积越多。
                if (_swallowed.Count > 0)
                {
                    List<int> dead = null;
                    foreach (var kv in _swallowed)
                        if (now - kv.Value > SwallowPairMs) (dead ??= new List<int>()).Add(kv.Key);
                    if (dead != null) foreach (int k in dead) _swallowed.Remove(k);
                }
                _swallowed[vk] = now;
            }
            return (IntPtr)1;
        }

        /// <summary>
        /// "吞掉这个键 → 去 UI 线程干事 → 干不成再把键原样还回去"。
        /// <para>全项目同一条铁律的**唯一**落地点，各分支不再各写一遍
        /// <c>Post(() =&gt; { if (!X()) SendVk(...); }); return Swallow(vk);</c>，
        /// 写漏一次就是"吞了但什么都没发生"——用户看到的是"快捷键全没了"。</para>
        /// </summary>
        /// <param name="vk">用户按的那个键（失败时按 <paramref name="replayOnFail"/> 决定要不要重放）。</param>
        /// <param name="restoreCtrl">重放时是否带上 Ctrl（用户手指还按着，SendVk 不会合成）。</param>
        /// <param name="restoreShift">重放时是否带上 Shift。</param>
        /// <param name="work">真正要做的动作；返回 false = 没做成。</param>
        /// <param name="hidePopup">做动作前先收起候选窗（改文本的动作都要）。</param>
        /// <param name="replayOnFail">
        /// 失败时要不要把键原样送回去。
        /// <para>**字母键与回车必须传 false**。SendVk 见到"Ctrl 正被物理按住"
        /// 会刻意不合成 Ctrl（见 Native.SendVk 的说明），于是"回放"出去的其实是
        /// 裸 `d` / `k` / 回车 —— 直接把字母打进用户代码、或者多插一行，
        /// 比"什么都没发生"糟得多。方向键这类**无副作用**的键才允许回放。</para>
        /// </param>
        private IntPtr SwallowAndRun(int vk, bool restoreCtrl, bool restoreShift,
            Func<bool> work, bool hidePopup = true, bool replayOnFail = true)
        {
            Post(() =>
            {
                if (hidePopup) _eng.Hide();
                if (work()) return;
                if (!replayOnFail)
                {
                    Log.Write("快捷键 vk=0x{0:X2} 没做成（键已吞掉，不回放裸字符）", vk);
                    return;
                }
                Native.SendVk(vk, restoreCtrl, restoreShift);
            });
            return Swallow(vk);
        }

        /// <summary>
        /// 钩子线程用：前台窗口是不是片段窗（或它的子控件）？
        /// 只读一个缓存的 HWND + `GetForegroundWindow` + 几次 `GetParent`，
        /// 全是纯 Win32，一次 COM 都不碰，也不去碰任何 WinForms 对象。
        /// </summary>
        private bool SnippetFormActive()
        {
            IntPtr fg = _snippetHwnd;
            if (fg == IntPtr.Zero) return false;
            IntPtr h = Native.GetForegroundWindow();
            for (int i = 0; i < 8 && h != IntPtr.Zero; i++)
            {
                if (h == fg) return true;
                IntPtr p = Native.GetParent(h);
                if (p == h) break;
                h = p;
            }
            return false;
        }

        /// <summary>钩子线程用：拿"轮询留下的快照"。快照过期 / 输入法可能正在组字时一律
        /// 返回 false —— 宁可这一次不接管（原生行为无害），也绝不拿旧行文本去赌
        /// 用户正在敲的那一行。
        /// </summary>
        private bool FreshHookSnap(out string lineText, out int caretDisp)
        {
            lineText = "";
            caretDisp = 0;
            var s = _hookSnap;
            if (s == null) return false;
            // 800ms：回车上下文的最大存活时间。在代码窗里打字时
            // 轮询是 50ms 一轮，这个上限其实是"轮询卡住了 / 刚切过焦点"的兜底。
            if (Environment.TickCount64 - _hookSnapTicks > 800) return false;
            // 出参先填上（即便最后返回 false）—— 回车的"快照落后"那一档
            //   还要拿这行的文本去判"看起来像不像块头"，见 HookProc 里的说明。
            lineText = s.Text ?? "";
            caretDisp = s.CaretCol;
            // 计数超过 2 秒还挂着 ⇒ 不是"输入法在组字"，是那次变化我们没采到。
            // 不清掉的话回车缩进 / 退格接管会被永久挡死（宁可放过这一次）。
            if (_pendingKeys > 0 && Environment.TickCount64 - _pendingAt > 2000) _pendingKeys = 0;
            if (_pendingKeys > 0) return false;
            return caretDisp > 0;
        }

        /// <summary>Ctrl+Q：开 / 关片段窗（UI 线程）。开窗位置取【开窗之前】的光标。</summary>
        private void ToggleSnippets()
        {
            try
            {
                if (_snippetForm != null && _snippetForm.Visible)
                {
                    _snippetForm.Close_();
                    return;
                }
                if (_snippetForm == null)
                {
                    _snippetStore = new SnippetStore();
                    _snippetForm = new SnippetForm(_snippetStore, code => _be.InsertSnippet(code));
                    _snippetForm.OnClosedCallback = () =>
                    {
                        _snippetHwnd = IntPtr.Zero;    // 钩子线程的实时判据一并失效
                        try { _be.FocusVbe(); } catch { }
                    };
                    // 用户直接点标题栏的 × 时 FormClosing 会转成 Hide + OnClosed，
                    // 但句柄被系统回收时不会 —— 这里补一道，免得 _snippetHwnd 卡住
                    _snippetForm.FormClosed += (s, e) => _snippetHwnd = IntPtr.Zero;
                }
                _eng.Hide();

                // 位置必须在窗子显示【之前】取：一显示前台就变了，插入符矩形就没了
                int x, top, bottom;
                bool has = Native.CaretRect(out x, out top, out bottom);
                // 句柄在这里（UI 线程上）逼出来并存给钩子线程用；窗子还没建句柄
                //   的时候 `_snippetHwnd` 就是 0，钩子那条实时判据会退化成"不拦"，
                //   而那时前台还在 VBE 代码窗 —— 正是我们想要的行为。
                _ = _snippetForm.Handle;
                _snippetHwnd = _snippetForm.Handle;
                _snippetForm.OpenAtCaret(has, x, top, bottom);
            }
            catch (Exception e)
            {
                _snippetHwnd = IntPtr.Zero;
                Log.Write("片段窗打开失败：{0}", e.Message);
            }
        }

        /// <summary>
        /// 是不是"我们该管的"组合键（用于诊断日志，避免每天打的字都进日志）。
        /// 列表只跟钩子里**真吃键**的分支对齐：Ctrl+D / Ctrl+K / Ctrl+Q /
        /// Ctrl+/ 一族 / 回车 / 上下键 / 空格。（Ctrl+Shift+L 与 Ctrl+Shift+↑↓
        /// 分别于 、删除，这里也一并摘掉了。）
        /// </summary>
        private static bool IsWatchedHotkey(int vk) =>
            vk == 'D' || vk == 'K' || vk == 'Q' ||
            Native.IsSlashKey(vk) ||
            vk == Native.VK_RETURN || vk == Native.VK_UP || vk == Native.VK_DOWN ||
            vk == Native.VK_SPACE;

        /// <summary>
        /// 诊断用：Ctrl + **任何一个符号键**（OEM 键区 0xBA~0xDF 与小键盘除号）
        /// 没命中任何分支时也要留痕。
        /// 不能只记录 <see cref="IsWatchedHotkey"/> 里那几个 vk，
        ///   于是"用户按的其实是另一个 vk"这种情况日志里一个字都没有 ——
        ///   排查"快捷键失效"时最怕的正是这个盲区。
        /// </summary>
        private static bool IsHotkeyDiagVk(int vk) =>
            IsWatchedHotkey(vk) || (vk >= 0xBA && vk <= 0xDF);

        /// <summary>
        /// 候选窗"现在真的还能用"吗？
        /// <para>为什么要心跳而不是只看 <see cref="_popupVisible"/>：UI 线程一旦卡住
        /// （COM 阻塞、Excel 弹模态框、调试断点），轮询就停了，而这个标志还是 true。
        /// 那一整段时间里 ↑/↓/PgUp/PgDn/Esc/Tab 全被我们吞掉，屏幕上毫无反应 ——
        /// 用户的感觉是"键盘死了"。3 秒没有心跳就当窗子已经不在，把键还给系统。</para>
        /// </summary>
        private bool PopupAlive =>
            _popupVisible && Environment.TickCount64 - _popupBeat < 3000;

        private void Post(Action a)
        {
            try { _disp.BeginInvoke(a); }
            catch (Exception e) { Log.Write("投递动作失败：{0}", e.Message); }
        }

        // ------------------------------------------------------------------
        // 轮询
        // ------------------------------------------------------------------

        /// <summary>纯 Win32：焦点是不是在 VBE 的代码窗格里（0.02ms，不碰 COM）。</summary>
        private static bool InVbeCodePane()
        {
            IntPtr frame = Native.FindWindow(VbeFrameClass, null);
            if (frame == IntPtr.Zero) return false;
            uint pid;
            uint tid = Native.GetWindowThreadProcessId(frame, out pid);
            IntPtr focus = Native.FocusedHwndOf(tid);
            if (focus == IntPtr.Zero) return false;

            string cls = Native.ClassNameOf(focus);
            foreach (var p in VbePopupClasses)
                if (string.Equals(cls, p, StringComparison.OrdinalIgnoreCase)) return true;

            IntPtr h = focus;
            for (int i = 0; i < 12 && h != IntPtr.Zero; i++)
            {
                if (string.Equals(Native.ClassNameOf(h), CodePaneClass, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (h == frame) break;
                IntPtr p = Native.GetParent(h);
                if (p == h) break;
                h = p;
            }
            // 拿不准（类名不认识）：退回"前台窗口属于 VBE 进程"
            string fgCls = Native.ClassNameOf(Native.GetForegroundWindow());
            return string.Equals(fgCls, VbeFrameClass, StringComparison.OrdinalIgnoreCase);
        }

        private void PollOnce()
        {
            if (!_running) return;

            // 最近使用记录：延迟合并写盘（2 秒最多一次）。之前是"每接受一个候选就
            // 整文件重写一次"，家目录落在 OneDrive / 网络盘上时每次 Tab 都卡一下。
            try
            {
                if (Environment.TickCount64 - _recentFlushAt > 2000)
                {
                    _recent.Flush();
                    _recentFlushAt = Environment.TickCount64;
                }
            }
            catch (Exception e) { Log.Write("recent 写盘失败：{0}", e.Message); }

            try
            {
                bool focused = InVbeCodePane();
                _codePaneFocused = focused;
                // 心跳：只要 UI 线程还活着、轮询还在跑，候选窗的"可用性"就续期一次
                if (_popupVisible) _popupBeat = Environment.TickCount64;

                if (!focused)
                {
                    if (_popupVisible) _eng.Hide();
                    _lastSnap = null;
                    _hookSnap = null;
                    _lastModName = "";
                    // 焦点已经不在代码窗：那些"按了键却没落进文档"的计数从此无从
                    // 求证（钩子在这之后也不再累加），留着只会白挡 2 秒的回车/
                    // 退格。刚切回 VBE 时同样清零。
                    _pendingKeys = 0;
                    SetInterval(250);
                    return;
                }

                Snapshot snap = _be.Snapshot();
                if (snap == null)
                {
                    if (_popupVisible) _eng.Hide();
                    _hookSnap = null;
                    SetInterval(150);
                    return;
                }

                // VBE 自己的提示词列表一出现，我们立刻收 —— 不能只靠
                //   "下次触发时让位"。触发那一刻 VBE 的列表往往还没画出来（它等
                //   WM_TIMER 才 Show），两边会并排挂在屏幕上几百毫秒，用户看到的
                //   就是一个压着另一个。这里每轮（打字时 50ms）查一次，
                //   缓存 150ms，开销约 0。
                if (_popupVisible && Native.VbeMemberListVisible())
                {
                    _eng.Hide();
                    // 日志限速：每 2 秒最多一条，别把 log 刷满
                    long now = Environment.TickCount64;
                    if (now - _lastYieldLog > 2000)
                    {
                        _lastYieldLog = now;
                        Log.Write("让位：VBE 的成员列表已出现，收起我们的候选窗");
                    }
                }

                bool typing = Environment.TickCount64 - _lastKeyTs < 900;
                SetInterval(typing ? 50 : 150);

                // 给钩子线程留一份"最近的行 + 光标列"：回车缩进 / 退格那几条判据
                // 全靠它（钩子线程一次 COM 都不许碰）。
                _hookSnapTicks = Environment.TickCount64;
                _hookSnap = snap;

                bool changed = _lastSnap == null
                    || snap.Line != _lastSnap.Line
                    || !string.Equals(Rules.Rstrip(snap.Text), Rules.Rstrip(_lastSnap.Text), StringComparison.Ordinal)
                    || !string.Equals(snap.Module, _lastModName, StringComparison.OrdinalIgnoreCase);

                if (changed)
                {
                    // 文档变了 ⇒ "按了键却没落进文档"的计数清零（见 _pendingKeys 的说明）
                    _pendingKeys = 0;

                    // 换模块 = 不是打字，只是切了代码窗，不该弹
                    bool modSwitched = !string.Equals(snap.Module, _lastModName, StringComparison.OrdinalIgnoreCase)
                                       && _lastSnap != null;

                    // **行号变了 ⇒ 光标是被"搬"过去的，不是在打字。**
                    //   否则会出现"鼠标单击到某个变量名就自动弹提示框"：
                    //   鼠标点到另一行上的变量时，快照的"行号 + 行文本"全变了，
                    //   看起来和敲了一个字一模一样；而"最近 1.2 秒按过键"拦不住
                    //   它（用户常常打完字顺手点一下变量看定义）。方向键上下 /
                    //   Ctrl+Home / 点滚动条也一样。
                    //   真打字时行号绝不会变 —— 回车换到的是空行，那种位置本来
                    //   也没词可提示，所以这条不会误伤正常输入。
                    bool lineJumped = _lastSnap != null && snap.Line != _lastSnap.Line;
                    int prevLine = _lastSnap == null ? 0 : _lastSnap.Line;

                    _lastSnap = snap;
                    _lastModName = snap.Module;

                    if (!modSwitched)
                    {
                        _eng.InvalidatePool();
                        // 文本变了不等于"用户在打字"：鼠标右键粘贴、拖拽一段
                        //   代码进来、别的程序通过 COM 改了模块，都会走到这里。
                        //   那些场景一律只收起、不弹（REQUIRE_KEY_TRIGGER）。
                        //   判据是"最近 1.2 秒内按过键"，以及"输入法可能正在组字"。
                        bool keyTriggered = !RequireKeyTrigger
                                            || _pendingKeys > 0
                                            || Environment.TickCount64 - _lastKeyTs <= KeyTriggerWindowMs;

                        // 剪贴板按键（Ctrl+C/X/V、Shift+Insert）之后的那一次变化：
                        //   内容确实变了，但变的是"粘进来的一段"，不是"敲出来的一个词"。
                        //   "最近按过键"这条对 Ctrl+V 照样成立（它也是按键），于是复制
                        //   一个变量名粘到别处、光标一落在词尾，候选窗就自己弹出来。
                        //   标记只消费一次、且有过期时间：剪贴板是空的、粘贴根本没落地
                        //   时，不能让它一直压着后面的正常输入。
                        if (_clipAt != 0)
                        {
                            if (Environment.TickCount64 - _clipAt <= ClipGuardMs) keyTriggered = false;
                            _clipAt = 0;
                        }
                        if (!Rules.AutoTriggerWanted(true, modSwitched, keyTriggered, lineJumped))
                        {
                            if (_popupVisible) _eng.Hide();
                            // 只有"本来会弹、因为跳行才没弹"这一种记一笔（限速 2 秒）。
                            //   用户再报"点了变量弹窗"时，日志里能直接看出是漏判还是别的原因。
                            if (lineJumped && keyTriggered)
                            {
                                long now = Environment.TickCount64;
                                if (now - _lastJumpLog > 2000)
                                {
                                    _lastJumpLog = now;
                                    Log.Write("不弹：光标跳行 L{0}->L{1}（鼠标单击 / 方向键），不是打字",
                                        prevLine, snap.Line);
                                }
                            }
                        }
                        else
                        {
                            // 不要在这里先 Hide 一下再 Trigger：Hide 会真的把窗口
                            //   收起来，Trigger 再 Show —— 用户每敲一个字都看到候选窗
                            //   "闪一下"。Trigger 内部该收的时候自己会收（打全即收 /
                            //   空结果 / 让位）。
                            _eng.Trigger(false, snap);
                        }
                    }
                }
                else if (_popupVisible)
                {
                    // 文本没变、光标动了 —— 那是"移动光标"，不是打字。
                    // 不能无条件 `ShowPopupAtCaret()`（把窗子挪到新光标处），
                    //   于是**在同一行里点一下鼠标 / 按左右方向键 / 按 Home**，候选窗
                    //   就跟着光标跑、列表一直挂着。"打 sh4 弹出 Sheet4 之后
                    //   点别处，提示不消失还跟着鼠标走"就是这条。
                    //   修空格时绕过它（空格特判自己收窗），但病根在这里：
                    //   "移动光标"一律该收起本次补全，跟按 ←/→ 的处理保持一致。
                    //   光标真没动（轮询空转）就什么都不做，窗子留在原位。
                    if (_lastSnap == null || _lastSnap.CaretCol != snap.CaretCol)
                    {
                        if (_popupVisible)
                            Log.Write("收起候选窗：光标在行内移动（列 {0}->{1}），不是打字",
                                _lastSnap == null ? -1 : _lastSnap.CaretCol, snap.CaretCol);
                        _eng.Hide();
                    }
                    _lastSnap = snap;      // 让比较基准前进，别每轮都判成"动了"
                }
            }
            catch (Exception e)
            {
                Log.Write("轮询异常：{0}", e.Message);
            }
        }

        private void SetInterval(int ms)
        {
            if (_timer == null) return;
            if (_timer.Interval != ms) _timer.Interval = ms;
        }

        private void ShowPopupAtCaret()
        {
            int x, top, bottom;
            if (!Native.CaretRect(out x, out top, out bottom))
            {
                // 拿不到光标就收起 —— 绝不猜一个位置糊上去
                _eng.Hide();
                return;
            }
            // ShowAt 是"先定位、再显示"：分开写（Show 之后才 PlaceAt）会让窗口
            //   在屏幕左上角闪一下再跳过来。
            _popup.ShowAt(x, top, bottom);
        }

        // ------------------------------------------------------------------
        // 退出
        // ------------------------------------------------------------------

        /// <summary>
        /// 退出。幂等：托盘"退出"和 `Application.Run` 返回后都会调一次，
        /// 第二次再对已 Dispose 的窗子 Hide/Dispose 只会制造 ObjectDisposedException。
        /// </summary>
        public void Shutdown()
        {
            if (Interlocked.Exchange(ref _shutdownDone, 1) != 0) return;
            _running = false;
            try { _timer?.Stop(); _timer?.Dispose(); _timer = null; } catch { }
            try { _recent.Flush(); } catch { }
            try
            {
                if (_hookThreadId != 0)
                    Native.PostThreadMessage(_hookThreadId, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            }
            catch { }
            try { if (_hookHandle != IntPtr.Zero) Native.UnhookWindowsHookEx(_hookHandle); } catch { }
            try { _showEvent?.Dispose(); } catch { }
            try { _popup.Hide(); _popup.Dispose(); } catch { }
            try { _snippetForm?.Hide(); _snippetForm?.Dispose(); } catch { }
            try { _snippetStore?.Dispose(); } catch { }
            try
            {
                if (_tray != null) { _tray.Visible = false; _tray.Dispose(); _tray = null; }
            }
            catch { }
            Log.Write("已退出");
            Log.Flush();          // 队列里最后几行别丢（后台线程是后台线程，进程一退就没了）
            ExitThread();
        }
    }
}
