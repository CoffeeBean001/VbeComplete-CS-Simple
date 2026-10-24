// Win32 层：全局键盘钩子、按键注入、窗口/进程探测。
//
// 这是"钩子线程"能碰的全部东西 —— 只有纯 Win32 和本进程内存。
//   钩子回调里【绝不允许】出现任何 COM 调用：VBE 可能正忙，一次阻塞就会
//   让整个系统的键盘卡住（低级钩子在输入队列里串行执行）。

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace VbeComplete
{
    internal static class Native
    {
        // ---- 虚拟键码（只列本项目用到的） ----
        public const int VK_BACK = 0x08;
        public const int VK_TAB = 0x09;
        public const int VK_RETURN = 0x0D;
        public const int VK_SHIFT = 0x10;
        public const int VK_CONTROL = 0x11;
        public const int VK_MENU = 0x12;    // Alt
        public const int VK_ESCAPE = 0x1B;
        public const int VK_SPACE = 0x20;
        public const int VK_PRIOR = 0x21;
        public const int VK_NEXT = 0x22;
        public const int VK_END = 0x23;
        public const int VK_HOME = 0x24;
        public const int VK_LEFT = 0x25;
        public const int VK_UP = 0x26;
        public const int VK_RIGHT = 0x27;
        public const int VK_DOWN = 0x28;
        public const int VK_INSERT = 0x2D;
        public const int VK_DELETE = 0x2E;
        public const int VK_C = 0x43;
        public const int VK_V = 0x56;
        public const int VK_X = 0x58;
        public const int VK_LWIN = 0x5B;
        public const int VK_RWIN = 0x5C;
        public const int VK_DIVIDE = 0x6F;  // 小键盘 /
        public const int VK_OEM_2 = 0xBF;   // /

        /// <summary>
        /// 这个 vk 是不是"`/` 键"—— 主键盘与**小键盘除号**都算。
        /// 代码里只认 VK_OEM_2 的话，小键盘除号按下去就像没发生一样。
        /// 判定只此一处，别再散落到各处比 ==。
        /// </summary>
        public static bool IsSlashKey(int vk) => vk == VK_OEM_2 || vk == VK_DIVIDE;

        /// <summary>
        /// 这一下按键是不是剪贴板操作：Ctrl+C / Ctrl+X / Ctrl+V、Shift+Insert。
        /// 用来给"这一下造成的文档变化"打标记 —— 粘贴是"原样放进来"，
        /// 不该被当成"用户敲完了一个词"。
        /// </summary>
        public static bool IsClipboardChord(bool ctrl, bool shift, bool alt, int vk)
        {
            if (alt) return false;
            if (ctrl && (vk == VK_C || vk == VK_X || vk == VK_V)) return true;
            return shift && !ctrl && vk == VK_INSERT;
        }

        // ---- 钩子 ----
        public const int WH_KEYBOARD_LL = 13;
        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;
        public const int WM_SYSKEYDOWN = 0x0104;
        public const int WM_SYSKEYUP = 0x0105;

        // KBDLLHOOKSTRUCT.flags
        public const uint LLKHF_INJECTED = 0x10;

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn,
            IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        // ---- 消息循环 ----
        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int ptX;
            public int ptY;
        }

        [DllImport("user32.dll")]
        public static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PostThreadMessage(uint idThread, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        public const uint WM_QUIT = 0x0012;

        // ---- 按键状态 ----
        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        public static bool Down(int vk) => (GetAsyncKeyState(vk) & 0x8000) != 0;

        // ---- 窗口 / 进程 ----
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // ---- 抢前台（只有探针用）----
        // SetForegroundWindow 对"当前不在前台的进程"会被系统直接拒掉（返回 false），
        //   这是 Windows 防焦点窃取的规则。绕法是先把自己的输入队列 Attach 到
        //   前台窗口所在线程，借它的"前台资格"再 SetForegroundWindow —— 探针要
        //   把 VBE 抢到前台才能让键盘钩子真的路由按键（钩子第一道闸就是焦点）。
        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo,
            [MarshalAs(UnmanagedType.Bool)] bool fAttach);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        /// <summary>把某个窗口强行弄到前台（探针专用，返回是否成功）。</summary>
        public static bool ForceForeground(IntPtr h)
        {
            if (h == IntPtr.Zero) return false;
            IntPtr fg = GetForegroundWindow();
            if (fg == h) return true;
            uint tidFg = fg == IntPtr.Zero ? 0 : GetWindowThreadProcessId(fg, out _);
            uint tidMe = GetCurrentThreadId();
            bool attached = tidFg != 0 && tidFg != tidMe && AttachThreadInput(tidMe, tidFg, true);
            try
            {
                ShowWindow(h, 9);              // SW_RESTORE：最小化的也拉出来
                BringWindowToTop(h);
                SetForegroundWindow(h);
                SetFocus(h);
            }
            finally
            {
                if (attached) AttachThreadInput(tidMe, tidFg, false);
            }
            return GetForegroundWindow() == h;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
            public int Width => Right - Left;
            public int Height => Bottom - Top;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        public const uint SWP_NOACTIVATE = 0x0010;
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X, Y; }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorPos(out POINT lpPoint);

        // ---- 按键注入 ----
        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx, dy;
            public uint mouseData, dwFlags, time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public ushort wVk, wScan;
            public uint dwFlags, time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL, wParamH;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct INPUTUNION
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public uint type;
            public INPUTUNION u;
        }

        public const uint INPUT_KEYBOARD = 1;
        public const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        public const uint KEYEVENTF_KEYUP = 0x0002;
        public const uint KEYEVENTF_UNICODE = 0x0004;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        /// <summary>注入一个 Unicode 字符（绕过输入法，直接进 WM_CHAR）。</summary>
        public static bool SendUnicodeChar(char ch)
        {
            INPUT[] seq = new INPUT[2];
            seq[0].type = INPUT_KEYBOARD;
            seq[0].u.ki.wVk = 0;
            seq[0].u.ki.wScan = ch;
            seq[0].u.ki.dwFlags = KEYEVENTF_UNICODE;
            seq[1] = seq[0];
            seq[1].u.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
            return SendInput(2, seq, Marshal.SizeOf(typeof(INPUT))) == 2;
        }

        /// <summary>注入一个虚拟键（不做修饰键合成，调用方自己决定按不按修饰键）。</summary>
        public static bool SendVk(int vk, bool ctrl = false, bool shift = false, bool alt = false)
        {
            // 物理按住的修饰键不要合成。命令失败、把原键还回去时，用户的手还按在
            //   Ctrl 上；我们再注入一对 Ctrl down/up，逻辑键状态会被"up"清掉 ——
            //   之后 Ctrl+C/V/Z 全失灵，按 D 直接把字母打进代码，直到用户松手重按。
            //   物理按着就一直按着：只发目标键，修饰状态由真实键盘维持。
            ctrl = ctrl && !Down(VK_CONTROL);
            shift = shift && !Down(VK_SHIFT);
            alt = alt && !Down(VK_MENU);

            int n = 0;
            INPUT[] seq = new INPUT[8];

            if (ctrl) { seq[n++] = VkInput(VK_CONTROL, false); }
            if (shift) { seq[n++] = VkInput(VK_SHIFT, false); }
            if (alt) { seq[n++] = VkInput(VK_MENU, false); }
            seq[n++] = VkInput(vk, false);
            seq[n++] = VkInput(vk, true);
            if (alt) { seq[n++] = VkInput(VK_MENU, true); }
            if (shift) { seq[n++] = VkInput(VK_SHIFT, true); }
            if (ctrl) { seq[n++] = VkInput(VK_CONTROL, true); }

            INPUT[] exact = new INPUT[n];
            Array.Copy(seq, exact, n);
            return SendInput((uint)n, exact, Marshal.SizeOf(typeof(INPUT))) == (uint)n;
        }

        private static INPUT VkInput(int vk, bool up)
        {
            INPUT it = new INPUT();
            it.type = INPUT_KEYBOARD;
            it.u.ki.wVk = (ushort)vk;
            it.u.ki.wScan = 0;
            it.u.ki.dwFlags = up ? KEYEVENTF_KEYUP : 0;
            // 方向键 / Delete 等需要 EXTENDED 位，否则某些应用收不到
            if (vk == VK_UP || vk == VK_DOWN || vk == VK_LEFT || vk == VK_RIGHT ||
                vk == VK_DELETE || vk == VK_HOME || vk == VK_END ||
                vk == VK_PRIOR || vk == VK_NEXT)
            {
                it.u.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
            }
            return it;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern IntPtr GetParent(IntPtr hWnd);

        // ---- VBE 自己的提示窗（成员列表 / 参数信息）----
        //
        // 为什么要知道它们在哪儿：VBE 的参数信息窗就画在光标正下方，我们的候选窗
        //   默认也在那儿（光标底边 +2px）。两窗一重叠，形参签名正好被候选列表压住
        //   —— 用户要的是"把候选窗下移到签名提示的下方"。
        //
        // 判据只认类名 + 可见 + 尺寸正常：句柄值会被 Windows 回收复用，只凭
        //   "还记得这个句柄"就相信它，一旦复用就会出现"永远以为 VBE 在弹提示"。
        private static readonly string[] VbePopupClasses = { "NameListWndClass", "PopupTipWndClass" };

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        private static long _popupRectsAt;
        private static bool _hasPopupCache;
        private static RECT[] _popupRectsCache = Array.Empty<RECT>();
        private static bool _memberListVisible;

        /// <summary>
        /// VBE 自己的【成员列表】(NameListWndClass) 此刻是不是画在屏幕上。
        /// 是的话我们就得让位 —— 两个列表叠在一起，用户根本不知道在看哪一个。
        /// 参数信息窗 (PopupTipWndClass) 不算：那个只要求我们挪开，不要求消失。
        /// </summary>
        public static bool VbeMemberListVisible()
        {
            VbePopupRects();        // 顺带刷新缓存
            return _memberListVisible;
        }

        /// <summary>
        /// 当前可见的 VBE 提示窗屏幕矩形。拿不到就返回空数组 ——
        /// **宁可重叠（可接受），也不因为探测失败把窗挪到奇怪的角落**。
        /// 缓存 150ms：枚举顶层窗口约 1.75ms，每轮轮询都做不可取；
        /// 但也不能太久 —— VBE 的成员列表弹出来之后我们要尽快让位。
        /// </summary>
        public static RECT[] VbePopupRects()
        {
            long now = Environment.TickCount64;
            // 原来这里写的是 `_popupRectsAt = long.MinValue`，配上
            //   `now - _popupRectsAt < 150` —— 减法**溢出**成负数，永远小于 150，
            //   于是每次都当成"缓存还新鲜"，`_popupRectsCache` 一直是初始的空数组。
            //   后果不是"少让位一次"，而是让位判据**从头到尾没生效过**：
            //   VbeMemberListVisible() 永远 false、避让 VBE 参数窗永远不挪。
            //   改用显式的 _hasPopupCache 标记，不再拿哨兵值做算术。
            if (_hasPopupCache && now - _popupRectsAt < 150) return _popupRectsCache;

            var found = new List<RECT>(4);
            bool memberList = false;
            try
            {
                IntPtr frame = FindWindow("wndclass_desked_gsk", null);
                uint hostPid = 0;
                if (frame != IntPtr.Zero) GetWindowThreadProcessId(frame, out hostPid);

                EnumWindows((h, p) =>
                {
                    try
                    {
                        string cls = ClassNameOf(h);
                        if (Array.IndexOf(VbePopupClasses, cls) < 0) return true;   // 继续枚举
                        if (!IsWindowVisible(h)) return true;
                        if (hostPid != 0)
                        {
                            uint pid;
                            GetWindowThreadProcessId(h, out pid);
                            if (pid != hostPid) return true;
                        }
                        RECT r;
                        if (!GetWindowRect(h, out r)) return true;
                        if (r.Right - r.Left <= 0 || r.Bottom - r.Top <= 0) return true;
                        found.Add(r);
                        if (string.Equals(cls, "NameListWndClass", StringComparison.OrdinalIgnoreCase))
                            memberList = true;
                    }
                    catch { }
                    return true;
                }, IntPtr.Zero);
            }
            catch { }

            _popupRectsCache = found.ToArray();
            _memberListVisible = memberList;
            _popupRectsAt = now;
            _hasPopupCache = true;
            return _popupRectsCache;
        }

        /// <summary>两个矩形是否真的重叠（只共边/共点不算）。</summary>
        public static bool Overlap(RECT a, RECT b) =>
            !(a.Right <= b.Left || a.Left >= b.Right || a.Bottom <= b.Top || a.Top >= b.Bottom);

        // ---- 线程焦点 / 插入符矩形 ----
        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO pgui);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        /// <summary>线程里当前拥有键盘焦点的那个窗口（拿不到返回 Zero）。</summary>
        public static IntPtr FocusedHwndOf(uint tid)
        {
            if (tid == 0) return IntPtr.Zero;
            var gti = new GUITHREADINFO();
            gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
            if (!GetGUIThreadInfo(tid, ref gti)) return IntPtr.Zero;
            return gti.hwndFocus;
        }

        /// <summary>VBE 线程里当前拥有键盘焦点的那个窗口（没有 VBE 就是 Zero）。</summary>
        public static IntPtr FocusedVbeWindow()
        {
            IntPtr frame = FindWindow("wndclass_desked_gsk", null);
            if (frame == IntPtr.Zero) return IntPtr.Zero;
            uint pid;
            uint tid = GetWindowThreadProcessId(frame, out pid);
            return FocusedHwndOf(tid);
        }

        // ---- 输入法组字状态 ----
        // 判"输入法正在组字"不能用"按了键但文档没变"这个计数，
        //   可它同时也把"刚敲完最后一个字符就按回车"（快照还没追上）算进去了 ——
        //   于是回车被放行，结构收尾补不出来。
        //   直接问输入法：有组字串 = 回车是【上屏键】，绝不能抢；没有 = 放心接管。
        [DllImport("imm32.dll")]
        private static extern IntPtr ImmGetContext(IntPtr hWnd);

        [DllImport("imm32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);

        [DllImport("imm32.dll", CharSet = CharSet.Unicode)]
        private static extern int ImmGetCompositionStringW(IntPtr hIMC, int dwIndex,
            byte[] lpBuf, int dwBufLen);

        private const int GCS_COMPSTR = 0x0008;

        /// <summary>这个窗口此刻有没有输入法组字串？没有/拿不到一律 false。</summary>
        public static bool ImeComposing(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return false;
            try
            {
                IntPtr himc = ImmGetContext(hwnd);
                if (himc == IntPtr.Zero) return false;
                try { return ImmGetCompositionStringW(himc, GCS_COMPSTR, null, 0) > 0; }
                finally { ImmReleaseContext(hwnd, himc); }
            }
            catch { return false; }
        }

        /// <summary>当前插入符的屏幕矩形：(x, top, bottom)。VBE 是自绘的，但它仍然
        /// 维护一个真的 caret（GetGUIThreadInfo.rcCaret），这是唯一可靠的位置来源。</summary>
        public static bool CaretRect(out int x, out int top, out int bottom)
        {
            x = top = bottom = 0;
            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero) return false;
            uint pid;
            uint tid = GetWindowThreadProcessId(fg, out pid);
            if (tid == 0) return false;

            var gti = new GUITHREADINFO();
            gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
            if (!GetGUIThreadInfo(tid, ref gti)) return false;
            if (gti.hwndCaret == IntPtr.Zero) return false;

            POINT p1 = new POINT { X = gti.rcCaret.Left, Y = gti.rcCaret.Top };
            POINT p2 = new POINT { X = gti.rcCaret.Left, Y = gti.rcCaret.Bottom };
            if (!ClientToScreen(gti.hwndCaret, ref p1)) return false;
            if (!ClientToScreen(gti.hwndCaret, ref p2)) return false;
            x = p1.X; top = p1.Y; bottom = p2.Y;
            return true;
        }

        public static string ClassNameOf(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return "";
            StringBuilder sb = new StringBuilder(256);
            int n = GetClassName(hwnd, sb, sb.Capacity);
            return n > 0 ? sb.ToString() : "";
        }

        public static uint PidOf(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return 0;
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            return pid;
        }

        public static string ProcessNameOfPid(uint pid)
        {
            if (pid == 0) return "";
            try
            {
                using (var p = System.Diagnostics.Process.GetProcessById((int)pid))
                    return p.ProcessName;
            }
            catch { return ""; }
        }

        // ------------------------------------------------------------------
        // 输入法：中文标点 / 全角时，标点不能被我们"翻成"半角
        //
        // 症状："我不能输入中文的双括号和中文双
        // 引号，一输入就会自动变成英文的双括号和英文双引号。"
        // 根因是自动配对只看"虚拟键 + Shift"，不看输入法状态 —— 中文标点档下
        // Shift+9 敲出来的是 `（`，我们却一律插半角 `()`。
        //
        // 取法（实测有效）：跨进程拿不到 ImmGetContext，只能
        // 走 `ImmGetDefaultIMEWnd(hwnd)` + `SendMessageTimeout(WM_IME_CONTROL,
        // IMC_GETCONVERSIONMODE)`。必须用 SendMessageTimeout + SMTO_ABORTIFHUNG：
        // 目标窗口卡死时立刻返回，绝不解把钩子线程钉死（低级键盘钩子超时会被
        // 系统悄悄摘掉，那是最坏的结果）。
        // ------------------------------------------------------------------
        private const uint WM_IME_CONTROL = 0x0283;
        private const int IMC_GETCONVERSIONMODE = 0x0001;
        private const int IMC_SETCONVERSIONMODE = 0x0002;
        private const int IME_CMODE_NATIVE = 0x0001;
        private const int IME_CMODE_FULLSHAPE = 0x0008;
        private const int IME_CMODE_SYMBOL = 0x0400;
        private const uint SMTO_ABORTIFHUNG = 0x0002;

        [DllImport("imm32.dll")]
        private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam,
            IntPtr lParam, uint flags, uint timeout, out IntPtr result);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        /// <summary>问一次窗口所属输入法的转换模式；问不到返回 -1（纯 Win32）。</summary>
        public static int ImeConversionMode(IntPtr hwnd)
        {
            try
            {
                if (hwnd == IntPtr.Zero) return -1;
                IntPtr ime = ImmGetDefaultIMEWnd(hwnd);
                if (ime == IntPtr.Zero) return -1;
                IntPtr res;
                IntPtr got = SendMessageTimeout(ime, WM_IME_CONTROL,
                    (IntPtr)IMC_GETCONVERSIONMODE, IntPtr.Zero, SMTO_ABORTIFHUNG, 50, out res);
                if (got == IntPtr.Zero) return -1;
                return (int)(res.ToInt64() & 0xFFFF);
            }
            catch { return -1; }
        }

        /// <summary>此刻敲出来的标点是不是全角中文标点（必须在中文模式 + 中文标点/全角）。</summary>
        public static bool ImeFullwidthPunct(IntPtr hwnd)
        {
            if (Settings.Flag("NO_IME_PUNCT", false)) return false;
            int m = ImeConversionMode(hwnd);
            if (m < 0) return false;
            return (m & IME_CMODE_NATIVE) != 0 &&
                   (m & (IME_CMODE_SYMBOL | IME_CMODE_FULLSHAPE)) != 0;
        }

        /// <summary>
        /// 输入法现在是不是"英数/半角"状态（<c>IME_CMODE_NATIVE</c> 没置位）。
        /// 只给真机探针用：打字类用例在中文模式下，每个字母都会被拿去组字，
        ///   文本根本落不进文档 —— 那不是程序坏了，是环境不对，探针必须先说清楚。
        ///   问不到（没有输入法 / 系统不答）按 true 处理，别让环境探测拖死用例。
        /// </summary>
        public static bool ImeAlphanumeric(IntPtr hwnd)
        {
            int m = ImeConversionMode(hwnd);
            if (m < 0) return true;
            return (m & IME_CMODE_NATIVE) == 0;
        }

        /// <summary>
        /// 把窗口所属输入法切到"英数"档（只清 <c>IME_CMODE_NATIVE</c>，别的位不动）。
        /// 只给真机探针用：探针要"逐字符打字"才能验回车接管，中文模式下字母会被
        ///   拿去组字（敲出来是 `Fori= 1 To3`），用例必然失败且原因假象为"程序坏了"。
        ///   走的是 IME 窗口的 <c>WM_IME_CONTROL / IMC_SETCONVERSIONMODE</c>，
        ///   与读取同一个通道、同样带 SMTO_ABORTIFHUNG —— 目标卡死也不会钉住我们。
        ///   返回切完之后的模式（-1 = 没问到）。
        /// </summary>
        public static int ImeSetAlphanumeric(IntPtr hwnd)
        {
            try
            {
                int m = ImeConversionMode(hwnd);
                if (m < 0) return -1;
                if ((m & IME_CMODE_NATIVE) == 0) return m;
                IntPtr ime = ImmGetDefaultIMEWnd(hwnd);
                if (ime == IntPtr.Zero) return m;
                IntPtr res;
                SendMessageTimeout(ime, WM_IME_CONTROL,
                    (IntPtr)IMC_SETCONVERSIONMODE, (IntPtr)(m & ~IME_CMODE_NATIVE),
                    SMTO_ABORTIFHUNG, 300, out res);
                return ImeConversionMode(hwnd);
            }
            catch { return -1; }
        }

        /// <summary>把输入法恢复成给定模式（探针跑完要把用户的输入法状态还回去）。</summary>
        public static void ImeRestore(IntPtr hwnd, int mode)
        {
            if (mode < 0 || hwnd == IntPtr.Zero) return;
            try
            {
                IntPtr ime = ImmGetDefaultIMEWnd(hwnd);
                if (ime == IntPtr.Zero) return;
                IntPtr res;
                SendMessageTimeout(ime, WM_IME_CONTROL,
                    (IntPtr)IMC_SETCONVERSIONMODE, (IntPtr)mode, SMTO_ABORTIFHUNG, 300, out res);
            }
            catch { }
        }

        /// <summary>
        /// 这个键敲出来的是不是需要自动配对的符号 —— 返回半角字符，否则 '\0'。
        /// 刻意去问系统"这个 vk 在当前键盘布局下的基础字符"（MapVirtualKey mode 2），
        /// 而不是写死 vk 常量：布局不同也不会配错。mode 0 返回的是扫描码不是字符，
        /// 别换成它。
        /// </summary>
        public static char PairCharForKey(int vk, bool shift)
        {
            if (!shift) return '\0';
            if (vk >= 0x60 && vk <= 0x69) return '\0';      // 小键盘不管
            try
            {
                uint baseCh = MapVirtualKey((uint)vk, 2) & 0xFFFF;
                if (baseCh == 0) return '\0';
                char ch = (char)baseCh;
                if (ch == '9') return '(';
                if (ch == '0') return ')';                 // 右半边已在右边时"跨过去"
                if (ch == '\'' || ch == '"') return '"';
                return '\0';
            }
            catch { return '\0'; }
        }

        /// <summary>半角配对字符 -> 输入法此刻真正会打出来的那个（中文标点档换全角）。</summary>
        public static char PairCharForIme(char half, IntPtr hwnd)
        {
            if (half == '(') return ImeFullwidthPunct(hwnd) ? '（' : half;
            // 右括号必须一起映射：少了这一条，中文标点档下自动配出的 `（` 按
            //   `)` 会走"跨过去"判定失败 → 注入半角 `)`，屏幕上是 `（)）`。
            if (half == ')') return ImeFullwidthPunct(hwnd) ? '）' : half;
            if (half == '"') return ImeFullwidthPunct(hwnd) ? '“' : half;
            return half;
        }

    }
}
