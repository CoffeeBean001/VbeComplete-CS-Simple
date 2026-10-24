// 候选窗：无边框、置顶、【永不抢焦点】、全部自绘。
//
// 四条不能动的规矩：
//   1. 绝不抢焦点（WS_EX_NOACTIVATE + ShowWithoutActivation）—— 抢一下焦点，
//      用户在 VBE 里的输入就断了，整个工具等于废掉；
//   2. 全部自绘（一个 OnPaint 画完）—— 每行一个 Control 的话，滚一下要重排
//      几十个句柄，那才是"卡"的来源；
//   3. 拿不到光标矩形就【不显示】，绝不猜一个位置糊上去；
//   4. 宽度按内容定（跟 VBE 的成员列表一样"够用就好"），上限 640px。
//
// 外观向 VBE 自带列表看齐：
//   * 字体 = 编辑器字体、**比 VBE 小一号**（14pt 的编辑器配 13pt 列表，
//     一样大反而显得"顶"；`POPUP_FONT_SIZE` 可改）；
//   * 白底黑字，命中的字母只**加粗**，不再涂红（VBE 只上黑字）；
//   * 行高 = 字高 + 4；
//   * 边框 1px 立体：上/左 #F0F0F0，下/右 #646464（截图逐像素取的色）；
//   * 滚动条用经典 3D 配色（#F0F0F0 槽 + #C0C0C0 滑块）。
//   保留的自家东西：图标（VBE 也有方法/属性图标）、点击选词、滚轮、详情行
//   （只在名字超过 640px 被截断时出现）。
//
// 两条界面口径：
//   1. 【字体】直接读 VBE 当前编辑器用的字体（`VBA\7.1` → `7.0` → `6.0` 依次找），
//      并去掉 Light/Thin 这类字重后缀 —— 细体放到候选窗里又小又扁。
//      都读不到时从"好看且等宽"的候选链里挑一个，不退回 Courier New。
//   2. 【选中行】VBE 选中文本那种**深蓝底 + 白字**，命中的字母在蓝底上用亮黄
//      （黑的加粗字压在深蓝上完全看不见）。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VbeComplete
{
    internal sealed class PopupForm : Form
    {
        // ---- 配色按 VBE 成员列表截图逐像素取值） ----
        private static readonly Color CBg = ColorTranslator.FromHtml("#FFFFFF");
        private static readonly Color CFg = ColorTranslator.FromHtml("#000000");
        private static readonly Color CHoverBg = ColorTranslator.FromHtml("#F0F0F0");
        private static readonly Color CDetailBg = ColorTranslator.FromHtml("#F0F0F0");
        private static readonly Color CDetailFg = ColorTranslator.FromHtml("#333333");
        private static readonly Color CDetailSep = ColorTranslator.FromHtml("#A0A0A0");
        private static readonly Color CTrack = ColorTranslator.FromHtml("#F0F0F0");
        private static readonly Color CThumb = ColorTranslator.FromHtml("#C0C0C0");
        // 立体边框：上/左取浅色、下/右取深色（VBE 的 NameListWndClass 就是这个配色）
        private static readonly Color CEdgeLight = ColorTranslator.FromHtml("#F0F0F0");
        private static readonly Color CEdgeDark = ColorTranslator.FromHtml("#646464");
        private static readonly Color CScrollEdge = ColorTranslator.FromHtml("#808080");

        // 选中行 = VBE 里"选中文本"那种深蓝（不是 Win10 那种浅蓝 #0078D7，
        //   那个在白色列表里偏亮、偏"系统对话框"，压不住 JetBrains Mono 的字形）。
        private static readonly Color CSelBg = SettingColor("POPUP_SEL_BG", "#0A62B0");
        private static readonly Color CSelFg = SettingColor("POPUP_SEL_FG", "#FFFFFF");
        // 蓝底上不能用黑字加粗标命中（看不见），改用亮黄 —— 对比度高且不刺眼
        private static readonly Color CSelHit = SettingColor("POPUP_SEL_HIT", "#FFD97A");

        /// <summary>读一个 #RRGGBB 形式的配色；写错 / 没写就用默认值，绝不让配色配置崩掉窗口。</summary>
        private static Color SettingColor(string key, string def)
        {
            string raw = Settings.Raw(key);
            foreach (var s in new[] { raw, def })
            {
                if (string.IsNullOrWhiteSpace(s)) continue;
                try
                {
                    var c = ColorTranslator.FromHtml(s.Trim());
                    if (c.A > 0) return c;
                }
                catch { }
            }
            return Color.Black;
        }

        // "好看且等宽"的候选链（从前到后第一个装了的赢）。等宽是关键：
        // 标识符列表里参差不齐的宽度看着比 VBE 还乱。
        private static readonly string[] PrettyFaces =
        {
            "JetBrains Mono", "Cascadia Mono", "Cascadia Code", "Consolas",
            "Source Code Pro", "Fira Code", "DejaVu Sans Mono",
            "Microsoft YaHei UI", "Segoe UI",
        };

        // 这些"出厂字体"又细又扁，VBE 拿它当默认不代表我们也该跟（Courier New 就是
        // 之前那个"扁平"的元凶）。
        private static readonly string[] UglyFaces =
        {
            "Courier New", "Courier", "Fixedsys", "Terminal", "SimSun", "宋体",
            "NSimSun", "新宋体", "MS Sans Serif", "Microsoft Sans Serif",
        };

        // 字体名里可以去掉的字重/字形后缀（"JetBrains Mono Light" -> "JetBrains Mono"）
        private static readonly string[] WeightSuffix =
        {
            "extra light", "extralight", "semi light", "semilight",
            "light", "thin", "medium", "semibold", "semi bold", "bold", "italic",
            "regular", "book",
        };

        private static readonly Dictionary<string, Color> IconColor =
            new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
            {
                { "func", ColorTranslator.FromHtml("#1F8A3B") },
                { "prop", ColorTranslator.FromHtml("#7A3FA8") },
                { "var",  ColorTranslator.FromHtml("#2A5CAA") },
                { "type", ColorTranslator.FromHtml("#B26A00") },
                { "kw",   ColorTranslator.FromHtml("#707070") },
            };

        private const int PadX = 2;
        private const int IconGap = 3;
        private const int ScrollW = 12;
        private const int MinChars = 10;         // 太窄会看不清（也免得窗口只有图标宽）
        private const string Ellipsis = "...";
        private const int MaxWidth = 640;
        private const int PopupGap = 2;          // 避让 VBE 提示窗时的留白
        private const int PopupTryMax = 4;

        private IReadOnlyList<Candidate> _items = Array.Empty<Candidate>();
        private int _sel;
        private int _top;
        private int _rows = 8;
        private int _rowH = 17;
        private int _iconW = 14;
        private int _hover = -1;
        private int _lastX = int.MinValue, _lastY = int.MinValue, _lastW, _lastH;

        private bool _dragScroll;
        private int _dragOffset;

        private readonly Font _font;
        private readonly Font _fontBold;
        private int _charW;          // 一个字符的像素宽（按 "x"*16 量）
        private int _minTextW;       // 至少留这么多字的宽度
        private Dictionary<string, int> _widthCache =
            new Dictionary<string, int>(StringComparer.Ordinal);

        public event Action<int> RowClicked;      // 单击某行（index 为绝对下标）
        public event Action<int> Wheel;           // 滚轮行数
        public event Action<int> ScrollTo;        // 滚动条拖动：目标 top

        public PopupForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = CBg;
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

            _font = BuildFont(out float pt);
            try
            {
                _fontBold = new Font(_font, FontStyle.Bold);
            }
            catch
            {
                _fontBold = _font;
            }
            _rowH = Math.Max(16, _font.Height + 4);
            _iconW = Math.Max(12, Math.Min(16, (int)(_font.Height * 0.8f)));
            // 原来写成 MeasureText("x", _font, 16) / 16 —— 量的是**一个**
            //   字符的宽度再除以 16，结果恒 < 1，被 Math.Max(4, …) 兜成 4（真实
            //   等宽字约 8~9px）。于是"最少 10 个字宽"的最窄宽度失效、Truncate
            //   高估一倍还要靠逐字符回退。量 16 个字符再除，才是本来想做的事。
            _charW = Math.Max(4, MeasureText(new string('x', 16), _font, 0) / 16);
            _minTextW = _charW * MinChars;

            TabStop = false;
        }

        /// <summary>
        /// 字体 = "好看 + 比 VBE 小一号"。
        ///   * 名字：`POPUP_FONT` 优先；否则用 VBE 编辑器字体（去掉 Light/Thin 字重）；
        ///     再不行从 <see cref="PrettyFaces"/> 里挑第一个装了的；
        ///   * 字号：`POPUP_FONT_SIZE` 正数 = 直接用它，负数 = 比 VBE 小几号，
        ///     0/没写 = 比 VBE 小一号（一样大反而显得"顶"）。
        ///     注册表读不到就退回 Courier New 10。
        /// </summary>
        private static Font BuildFont(out float pt)
        {
            float vbePt;
            string vbeFace = ReadVbeFont(out vbePt);

            string face = Settings.Raw("POPUP_FONT");
            double want = Settings.Double("POPUP_FONT_SIZE", 0);

            if (want > 0) pt = (float)want;
            else if (want < 0) pt = vbePt + (float)want;
            else pt = vbePt - 1;            // 默认比 VBE 小一号

            if (pt <= 0) pt = 10;
            if (pt < 8) pt = 8;
            if (pt > 24) pt = 24;           // 上限 24pt，再大就压住代码了

            if (string.IsNullOrWhiteSpace(face)) face = PickFace(vbeFace);
            return CreateFont(face.Trim(), pt);
        }

        /// <summary>
        /// 读 VBE 编辑器当前的字体与字号。按 7.1 → 7.0 → 6.0 的顺序找 ——
        /// 只写 7.1 的机器上停在 6.0 就永远读不到，会静默退回 Courier New 10，
        /// 候选窗看起来又小又扁。
        /// </summary>
        private static string ReadVbeFont(out float pt)
        {
            pt = 0;
            string face = null;
            foreach (var ver in new[] { "7.1", "7.0", "6.0" })
            {
                try
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\" + ver + @"\Common"))
                    {
                        if (k == null) continue;
                        if (string.IsNullOrEmpty(face)) face = k.GetValue("FontFace") as string;
                        if (pt <= 0)
                        {
                            object h = k.GetValue("FontHeight");
                            // FontHeight 是**像素**，不是点。直接 Convert.ToSingle 当点用的话，
                            //   96dpi 下等于放大 4/3 —— 候选窗字体就会比 VBE 编辑器大一圈。
                            //   这里换算成点再用。
                            try { if (h != null) pt = PxToPt(Convert.ToSingle(h)); } catch { }
                        }
                        if (!string.IsNullOrEmpty(face) && pt > 0) break;
                    }
                }
                catch { }
            }
            return face;
        }

        /// <summary>
        /// 像素 -> 点。VBE 注册表里的 <c>FontHeight</c> 是像素（本机 14），
        /// 而 Font 构造只要点。96dpi 下 1pt = 4/3 px，直接当点用会大 33%。
        /// 按屏幕 DPI 换算，取不到就按 96 算（px*72/96 兜底）。
        /// </summary>
        private static float PxToPt(float px)
        {
            float dpi = 96f;
            try
            {
                using (var g = Graphics.FromHwnd(IntPtr.Zero))
                    dpi = g.DpiY > 0 ? g.DpiY : 96f;
            }
            catch { }
            return px * 72f / dpi;
        }

        /// <summary>
        /// 挑字体名。先试 VBE 编辑器字体（去掉字重后缀 —— 用户设的
        /// "JetBrains Mono Light" 是细体，小字号下正是"扁平"的来源），
        /// 再退到 <see cref="PrettyFaces"/> 候选链。
        /// </summary>
        private static string PickFace(string vbeFace)
        {
            if (!string.IsNullOrEmpty(vbeFace))
            {
                string raw = vbeFace.Trim();
                // 仍然**先去字重**（"JetBrains Mono Light" -> "JetBrains Mono"）。
                //   试过优先用原样族名，结果 GDI+ 对 "… Light" 这种名字时灵时不灵
                //   （bin 版解析成 JetBrains Mono、发布版解析成 Light），一旦解析不到
                //   就会 fallback 成**比例字体** —— 候选窗里所有名字立刻歪掉。
                //   WinForms 的 Font 本来也只有 Regular/Bold/Italic 三档，拿不到 Light。
                string baseFace = StripWeight(raw);
                if (!IsUgly(baseFace) && FamilyExists(baseFace)) return baseFace;
                if (!IsUgly(raw) && FamilyExists(raw)) return raw;
            }
            foreach (var f in PrettyFaces) if (FamilyExists(f)) return f;
            return "Courier New";
        }

        private static string StripWeight(string face)
        {
            foreach (var w in WeightSuffix)
            {
                if (!face.EndsWith(w, StringComparison.OrdinalIgnoreCase)) continue;
                int cut = face.Length - w.Length;
                if (cut <= 0) continue;
                char prev = face[cut - 1];
                if (prev != ' ' && prev != '-' && prev != '_') continue;
                string trimmed = face.Substring(0, cut).TrimEnd(' ', '-', '_');
                if (trimmed.Length > 0) return trimmed;
            }
            return face;
        }

        private static bool IsUgly(string face)
        {
            foreach (var u in UglyFaces)
                if (string.Equals(face, u, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        private static bool FamilyExists(string face)
        {
            if (string.IsNullOrWhiteSpace(face)) return false;
            try
            {
                using (var ff = new FontFamily(face)) { return true; }
            }
            catch { return false; }
        }

        /// <summary>建字体；装不上（GDI 会悄悄换成别的字）就退回 Courier New 同字号。</summary>
        private static Font CreateFont(string face, float pt)
        {
            try
            {
                var f = new Font(face, pt, FontStyle.Regular, GraphicsUnit.Point);
                if (string.Equals(f.Name, face, StringComparison.OrdinalIgnoreCase)) return f;
                f.Dispose();
            }
            catch { }
            return new Font("Courier New", pt, FontStyle.Regular, GraphicsUnit.Point);
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;   // WS_EX_NOACTIVATE：点了也不抢焦点
                cp.ExStyle |= 0x00000080;   // WS_EX_TOOLWINDOW：不出现在 Alt-Tab
                return cp;
            }
        }

        public IReadOnlyList<Candidate> Items => _items;
        public int RowHeight => _rowH;
        /// <summary>实际用上的字体（"JetBrains Mono 14"）—— 排障时一眼看出是不是没读到 VBE 的字号。</summary>
        public string FontName =>
            _font.Name + " " + _font.SizeInPoints.ToString("0.#") +
            (ReferenceEquals(_fontBold, _font) ? " (无粗体)" : " +bold");

        public void SetData(IReadOnlyList<Candidate> items, int sel, int top, int rows)
        {
            _items = items ?? Array.Empty<Candidate>();
            _sel = sel;
            _top = top;
            _rows = Math.Max(1, rows);
            _hover = -1;
            Invalidate();
        }

        public void UpdateView(int sel, int top)
        {
            _sel = sel;
            _top = top;
            Invalidate();
        }

        public int VisibleRows => Math.Max(1, Math.Min(_rows, _items.Count - _top));
        public int ListHeight => VisibleRows * _rowH;
        public int TotalHeight => ListHeight + DetailHeight;

        /// <summary>
        /// 窗口该有多大（离屏渲染与自检用；运行期由 PlaceAt 自己算）。
        /// 注意别叫 PreferredSize —— 那是 <see cref="Control"/> 上的属性,同名方法会隐藏它。
        /// </summary>
        public Size DesiredSize() => new Size(ComputeWidth(), TotalHeight + 2);

        /// <summary>当前该不该显示详情行：只有"活动行的名字被截断"时才显示。</summary>
        public string DetailText()
        {
            int active = _hover >= 0 ? _hover : _sel;
            if (active < 0 || active >= _items.Count) return null;
            string name = _items[active] == null ? null : _items[active].Name;
            if (string.IsNullOrEmpty(name)) return null;
            return MeasureText(name, _font, 0) > ContentWidth ? name : null;
        }

        public int DetailHeight
        {
            get { return DetailText() == null ? 0 : _font.Height + 4; }
        }

        private int ContentWidth => Math.Max(MeasureLongest(), _minTextW);

        /// <summary>
        /// 当前候选里最长的那个名字有多宽（按名字缓存）。**不许看 DetailText** ——
        /// 详情行的宽度又是按这里算的，互相引用就成了无限递归）。
        /// </summary>
        private int MeasureLongest()
        {
            int longest = 0;
            for (int i = 0; i < _items.Count; i++)
            {
                string name = _items[i] == null ? null : _items[i].Name;
                if (string.IsNullOrEmpty(name)) continue;
                int w;
                if (!_widthCache.TryGetValue(name, out w))
                {
                    w = MeasureText(name, _font, 0);
                    if (_widthCache.Count < 4096) _widthCache[name] = w;
                }
                if (w > longest) longest = w;
            }
            return longest;
        }

        /// <summary>
        /// 宽度 = 最长候选 + 图标 + 内边距 + 滚动条，至少 MinChars 个字符宽。
        /// 跟 VBE 一样"够用就好"：VBE 的列表也是贴着最长条目收窄的。
        /// 每个名字的宽度按名字缓存 —— 每次按键都会重算一次宽度，逐条 MeasureText
        /// 在几十条候选上会变成可感知的开销。
        /// </summary>
        private int ComputeWidth()
        {
            int w2 = PadX + _iconW + IconGap + ContentWidth + ScrollW + PadX;
            string detail = DetailText();
            if (detail != null)
            {
                int dw = PadX + MeasureText(detail, _font, 0) + PadX;
                if (dw > w2) w2 = dw;
            }
            return Math.Min(MaxWidth, Math.Max(120, w2));
        }

        /// <summary>
        /// 贴着光标显示。(x, top, bottom) 是插入符的屏幕坐标。
        /// 下方放不下就翻到光标上方；若与 VBE 的参数信息窗重叠，整体下移到它下方。
        /// </summary>
        public void PlaceAt(int x, int top, int bottom)
        {
            int w = ComputeWidth();
            int detailH = DetailHeight;
            int listH = VisibleRows * _rowH;
            int h = listH + detailH;

            var scr = Screen.FromPoint(new Point(x, Math.Max(0, bottom)));
            Rectangle wa = scr.WorkingArea;

            int px = x;
            int py = bottom + 2;
            if (py + h > wa.Bottom) py = top - h - 2;

            // 避让 VBE 的提示窗（参数信息 / 成员列表）：只在"水平也相交"时才躲。
            var rects = Native.VbePopupRects();
            if (rects.Length > 0)
            {
                for (int guard = 0; guard < PopupTryMax; guard++)
                {
                    bool moved = false;
                    foreach (var r in rects)
                    {
                        var me = new Native.RECT { Left = px, Top = py, Right = px + w, Bottom = py + h };
                        if (!Native.Overlap(me, r)) continue;
                        py = r.Bottom + PopupGap;
                        moved = true;
                    }
                    if (!moved) break;
                }
                if (py + h > wa.Bottom)
                {
                    int alt = top - h - PopupGap;
                    bool clash = false;
                    foreach (var r in rects)
                    {
                        var me = new Native.RECT { Left = px, Top = alt, Right = px + w, Bottom = alt + h };
                        if (Native.Overlap(me, r)) { clash = true; break; }
                    }
                    py = (!clash && alt >= wa.Top) ? alt : Math.Max(wa.Top, wa.Bottom - h);
                }
            }

            if (py < wa.Top) py = wa.Top;
            if (px + w > wa.Right) px = wa.Right - w;
            if (px < wa.Left) px = wa.Left;

            // 位置没变就别再 SetWindowPos（每帧都调会让窗口轻微闪）
            if (_lastX == px && _lastY == py && _lastW == w && _lastH == h && Visible) return;
            _lastX = px; _lastY = py; _lastW = w; _lastH = h;

            Native.SetWindowPos(Handle, Native.HWND_TOPMOST, px, py, w, h, Native.SWP_NOACTIVATE);
            Width = w; Height = h;
        }

        public void ResetPlacement()
        {
            _lastX = int.MinValue; _lastY = int.MinValue;
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // 自绘，避免闪烁
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.SmoothingMode = SmoothingMode.None;      // 点线框 / 1px 边要锐利

            int visible = VisibleRows;
            int listH = visible * _rowH;
            string detail = DetailText();

            using (var b = new SolidBrush(CBg)) g.FillRectangle(b, 0, 0, Width, listH);
            if (detail != null)
                using (var b = new SolidBrush(CDetailBg)) g.FillRectangle(b, 0, listH, Width, Height - listH);

            // 选中条不能压在滚动条底下（跟 VBE 一样，滚动条是一条独立的长条）
            int barW = _items.Count > _rows ? ScrollW : 0;
            int rowW = Math.Max(1, Width - 1 - barW);

            for (int i = 0; i < visible; i++)
            {
                int idx = _top + i;
                if (idx >= _items.Count) break;
                var c = _items[idx];
                if (c == null) continue;
                int y = i * _rowH;
                bool selected = idx == _sel;

                if (selected)
                    using (var b = new SolidBrush(CSelBg)) g.FillRectangle(b, 1, y, rowW, _rowH);
                else if (idx == _hover)
                    using (var b = new SolidBrush(CHoverBg)) g.FillRectangle(b, 1, y, rowW, _rowH);

                int ix = PadX;
                int iy = y + (_rowH - _iconW) / 2;
                DrawIcon(g, c.Kind, ix, iy, _iconW, selected);

                int tx = ix + _iconW + IconGap;
                DrawName(g, c, tx, y, selected);
            }

            // 详情行：完整名字（列表里被截断时看这里；宽度按内容算之后极少出现）
            if (detail != null)
            {
                using (var pen = new Pen(CDetailSep)) g.DrawLine(pen, 0, listH, Width, listH);
                using (var b = new SolidBrush(CDetailFg))
                {
                    var rf = new RectangleF(PadX + 1, listH + (Height - listH - _font.Height) / 2f,
                        Width - PadX * 2, _font.Height + 2);
                    using (var sf = new StringFormat { Trimming = StringTrimming.EllipsisCharacter })
                        g.DrawString(detail, _font, b, rf, sf);
                }
            }

            // 滚动条（只有真需要时才画）
            if (_items.Count > _rows)
            {
                int trackX = Width - ScrollW;
                using (var b = new SolidBrush(CTrack)) g.FillRectangle(b, trackX, 0, ScrollW, listH);
                using (var pen = new Pen(CScrollEdge)) g.DrawLine(pen, trackX, 0, trackX, listH - 1);
                int total = _items.Count;
                int thumbH = Math.Max(16, (int)((long)listH * Math.Min(_rows, total) / total));
                int maxTop = Math.Max(1, total - _rows);
                int thumbY = (int)((long)(listH - thumbH) * Math.Min(_top, maxTop) / maxTop);
                var thumb = new Rectangle(trackX + 2, thumbY, ScrollW - 3, thumbH - 1);
                using (var b = new SolidBrush(CThumb)) g.FillRectangle(b, thumb);
                using (var pen = new Pen(Color.White)) g.DrawLine(pen, thumb.Left, thumb.Top, thumb.Right, thumb.Top);
                using (var pen = new Pen(CEdgeDark))
                {
                    g.DrawLine(pen, thumb.Left, thumb.Bottom - 1, thumb.Right, thumb.Bottom - 1);
                    g.DrawLine(pen, thumb.Right - 1, thumb.Top, thumb.Right - 1, thumb.Bottom - 1);
                }
            }

            // 1px 立体边：上/左浅、下/右深（与 VBE 列表一致）
            using (var pen = new Pen(CEdgeLight))
            {
                g.DrawLine(pen, 0, 0, Width - 1, 0);
                g.DrawLine(pen, 0, 0, 0, Height - 1);
            }
            using (var pen = new Pen(CEdgeDark))
            {
                g.DrawLine(pen, 0, Height - 1, Width - 1, Height - 1);
                g.DrawLine(pen, Width - 1, 0, Width - 1, Height - 1);
            }
        }

        /// <summary>
        /// 按"命中/未命中"把名字切成若干段再上色。
        /// 相邻同类字符合并成一段 —— 一个名字通常只有 2~4 段，绘制开销可以忽略。
        /// </summary>
        public static List<KeyValuePair<string, bool>> SplitByHits(string name, int[] positions)
        {
            var outp = new List<KeyValuePair<string, bool>>();
            if (string.IsNullOrEmpty(name)) return outp;
            var hit = new HashSet<int>();
            if (positions != null) foreach (var p in positions) hit.Add(p);

            int i = 0, n = name.Length;
            while (i < n)
            {
                bool isHit = hit.Contains(i);
                int j = i + 1;
                while (j < n && hit.Contains(j) == isHit) j++;
                outp.Add(new KeyValuePair<string, bool>(name.Substring(i, j - i), isHit));
                i = j;
            }
            return outp;
        }

        /// <summary>
        /// 画名字。命中的字母只**加粗**，不再涂红 —— VBE 的列表全是黑字，
        /// 涂色就不像了；加粗既保留"这几个字母命中了"的提示，又不改配色。
        /// 选中行是深蓝底，黑字加粗在上面完全看不见 —— 那里用白字 + 亮黄。
        /// </summary>
        private void DrawName(Graphics g, Candidate c, int x, int y, bool selected)
        {
            string name = c.Name ?? "";
            int baselineY = y + (_rowH - _font.Height) / 2;

            Color fg = selected ? CSelFg : CFg;
            Color hitFg = selected ? CSelHit : CFg;

            string disp = Truncate(name, ContentWidth, out int prefixLen);
            var hits = new List<int>();
            if (c.Hits != null)
                foreach (var p in c.Hits) if (p >= 0 && p < prefixLen) hits.Add(p);

            if (hits.Count == 0)
            {
                TextRenderer.DrawText(g, disp, _font, new Point(x, baselineY), fg,
                    TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
                return;
            }

            int cx = x;
            foreach (var seg in SplitByHits(disp, hits.ToArray()))
                cx = DrawChunk(g, seg.Key, cx, baselineY, seg.Value, fg, hitFg);
        }

        private int DrawChunk(Graphics g, string s, int x, int y, bool bold, Color fg, Color hitFg)
        {
            if (string.IsNullOrEmpty(s)) return x;
            var font = bold && !ReferenceEquals(_fontBold, _font) ? _fontBold : _font;
            TextRenderer.DrawText(g, s, font, new Point(x, y), bold ? hitFg : fg,
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            return x + MeasureText(s, font, 0);
        }

        private int MeasureText(string s, Font font, int maxW)
        {
            return MeasureText(s, font, maxW, _rowH);
        }

        private static int MeasureText(string s, Font font, int maxW, int h)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            var sz = TextRenderer.MeasureText(s, font,
                new Size(maxW > 0 ? maxW : int.MaxValue, Math.Max(8, h)),
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            return sz.Width;
        }

        /// <summary>
        /// 超过 maxW 就用 ... 省略。prefixLen 是"原名字里还能完整显示几个字符"
        /// （命中下标超出这个范围的不再上色 —— 它们已经被省略号吃掉了）。
        /// 先按等宽字体的字符宽估算再微调：逐字符 MeasureText 在长名字上是 O(n)
        /// 次 GDI 调用，而这一步每帧每行都要做。
        /// </summary>
        public string Truncate(string name, int maxW, out int prefixLen)
        {
            prefixLen = name == null ? 0 : name.Length;
            if (string.IsNullOrEmpty(name)) return "";
            if (MeasureText(name, _font, 0) <= maxW) return name;

            int budget = Math.Max(0, maxW - MeasureText(Ellipsis, _font, 0));
            int cut = Math.Min(name.Length, Math.Max(0, budget / Math.Max(1, _charW)));
            while (cut > 0 && MeasureText(name.Substring(0, cut), _font, 0) > budget) cut--;
            if (cut == 0 && name.Length > 0) cut = 1;

            prefixLen = cut;
            return name.Substring(0, cut) + Ellipsis;
        }

        /// <summary>
        /// 五档图标全部现画（不发位图）：函数=绿箭头、属性=紫方块、变量=蓝竖条、
        /// 类型=琥珀菱形、关键字=灰横线。**拿不到档位就不画**。
        /// 选中行整行是深蓝底：图标统一改白，否则蓝竖条会沉进背景里。
        /// </summary>
        private static void DrawIcon(Graphics g, string kind, int x, int y, int w, bool selected)
        {
            if (string.IsNullOrEmpty(kind)) return;
            Color color;
            if (!IconColor.TryGetValue(kind, out color)) return;
            if (selected) color = CSelFg;

            using (var b = new SolidBrush(color))
            using (var pen = new Pen(color))
            {
                switch (kind)
                {
                    case "func":   // 绿箭头（VBE 的方法图标就是一支箭头）
                        {
                            var pts = new[]
                            {
                                new Point(x, y + w / 4),
                                new Point(x + w / 2, y + w / 4),
                                new Point(x + w / 2, y),
                                new Point(x + w, y + w / 2),
                                new Point(x + w / 2, y + w),
                                new Point(x + w / 2, y + w * 3 / 4),
                                new Point(x, y + w * 3 / 4),
                            };
                            g.FillPolygon(b, pts);
                            break;
                        }
                    case "prop":   // 紫色实心方块
                        g.FillRectangle(b, x + 1, y + w / 4, w - 2, w / 2);
                        break;
                    case "type":   // 琥珀菱形
                        {
                            var pts = new[]
                            {
                                new Point(x + w / 2, y),
                                new Point(x + w, y + w / 2),
                                new Point(x + w / 2, y + w),
                                new Point(x, y + w / 2),
                            };
                            g.FillPolygon(b, pts);
                            break;
                        }
                    case "kw":     // 灰色横线
                        g.FillRectangle(b, x + 1, y + w / 2 - 1, w - 2, 2);
                        break;
                    default:       // var：蓝色小竖条
                        g.FillRectangle(b, x + w / 3, y + 1, Math.Max(2, w / 4), w - 2);
                        break;
                }
            }
        }

        // ------------------------------------------------------------------
        // 鼠标
        // ------------------------------------------------------------------

        private int RowAt(int y)
        {
            if (y < 0 || y >= ListHeight) return -1;
            int i = y / Math.Max(1, _rowH);
            int idx = _top + i;
            return idx < _items.Count ? idx : -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_dragScroll)
            {
                int listH = ListHeight;
                int total = _items.Count;
                int maxTop = Math.Max(0, total - _rows);
                if (maxTop > 0 && listH > 0)
                {
                    int thumbH = Math.Max(16, (int)((long)listH * Math.Min(_rows, total) / total));
                    int span = Math.Max(1, listH - thumbH);
                    int y = Math.Max(0, Math.Min(span, e.Y - _dragOffset));
                    int top = (int)((long)y * maxTop / span);
                    ScrollTo?.Invoke(top);
                }
                return;
            }

            // 滚动条区域不参与悬停高亮
            int hover = e.X >= Width - ScrollW && _items.Count > _rows ? -1 : RowAt(e.Y);
            if (hover != _hover)
            {
                _hover = hover;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            // 捕获期间一般不会收到 Leave；这里只是兜底复位，防止窗口被强制隐藏
            // 时把 _dragScroll 留在 true。
            _dragScroll = false;
            if (_hover != -1) { _hover = -1; Invalidate(); }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            int listH = ListHeight;
            if (e.X >= Width - ScrollW && _items.Count > _rows && e.Y < listH)
            {
                int total = _items.Count;
                int thumbH = Math.Max(16, (int)((long)listH * Math.Min(_rows, total) / total));
                int maxTop = Math.Max(1, total - _rows);
                int thumbY = (int)((long)(listH - thumbH) * Math.Min(_top, maxTop) / maxTop);
                if (e.Y >= thumbY && e.Y <= thumbY + thumbH)
                {
                    _dragScroll = true;
                    _dragOffset = e.Y - thumbY;
                    // 必须取鼠标捕获。否则把滑块拖出窗口再松手时收不到
                    //   MouseUp，_dragScroll 永远卡在 true —— 之后鼠标在窗口里
                    //   一动就继续滚，悬停高亮也再也不更新。
                    Capture = true;
                }
                else
                {
                    // 点空白轨道：翻一页
                    Wheel?.Invoke(e.Y < thumbY ? -PageRows : PageRows);
                }
                return;
            }

            int i = RowAt(e.Y);
            if (i >= 0) RowClicked?.Invoke(i);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _dragScroll = false;
            if (Capture) Capture = false;      // 与 OnMouseDown 的 Capture 成对
        }

        private int PageRows => Math.Max(1, _rows - 1);

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            Wheel?.Invoke(e.Delta > 0 ? -3 : 3);
        }

        /// <summary>
        /// 先定位、再显示。顺序反了用户会看到窗口在屏幕左上角闪一下再跳过来：
        /// <c>StartPosition = Manual</c> 且还没设过 Location 时，首次 Show 用的是
        /// 系统给的默认位置（0,0 附近）。
        /// </summary>
        public void ShowAt(int x, int top, int bottom)
        {
            PlaceAt(x, top, bottom);       // 内部会碰 Handle，句柄在这里建出来
            if (!Visible) Show();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // 字体是本类自己 new 的（BuildFont / new Font(_font, Bold)）——
                // 不释放就每个实例漏一个 HFONT。同一实例被 _fontBold 引用时只释放一次。
                if (!ReferenceEquals(_fontBold, _font)) _fontBold?.Dispose();
                _font?.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            const int WM_MOUSEACTIVATE = 0x0021;
            const int MA_NOACTIVATE = 3;
            if (m.Msg == WM_MOUSEACTIVATE)
            {
                m.Result = (IntPtr)MA_NOACTIVATE;
                return;
            }
            base.WndProc(ref m);
        }
    }
}
