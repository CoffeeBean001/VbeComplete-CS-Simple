// 代码片段窗（Ctrl+Q）。
//
// 这个窗口的视觉/布局是独立的，刻意与常见补全器长得不一样：
//     · 顶部工具栏："片段库" + 搜索框（占满剩余宽度，右边缘与主体对齐）
//     · 左侧自绘列表：每条两行 —— 标题（搜索命中逐字符高亮）+ 代码首行预览
//     · 右侧：标题输入框 + 代码编辑区（等宽字体、Tab 缩进、Ctrl+Enter 插入）
//     · 底部状态条 + 主次分明的按钮（"插入"是唯一的高亮主按钮）
//
// 为什么列表要自绘而不是用 ListBox：
//   1) 搜索命中要高亮到**逐字符**（与候选窗同一套观感）；
//   2) ListBox 的 DrawToBitmap 在离屏渲染时经常画不出条目 —— --uitest 就废了。
//   自绘一个 Control 之后，OnPaint 里画什么、截图里就是什么。
//
// 窗口形态与补全弹窗（PopupForm）刻意相反：补全窗是 NOACTIVATE 的浮窗、
//   一个键都不收；这个窗要收键盘（搜索框打字、代码框改代码），所以是普通窗口，
//   会抢焦点、可拖动、可缩放。代价是"打开期间焦点不在 VBE"，全局钩子的焦点
//   判据会把按键全部放行 —— 于是窗内的 Esc / Ctrl+Q 必须由本窗自己绑。
//
// 数据在 SnippetStore，插入动作由 App 注入（onInsert），本文件**不碰 COM**。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace VbeComplete
{
    internal sealed class SnippetForm : Form
    {
        // ---- 配色 ----
        private static readonly Color CBody = ColorTranslator.FromHtml("#F5F7FA");
        private static readonly Color CCard = ColorTranslator.FromHtml("#FFFFFF");
        private static readonly Color CLine = ColorTranslator.FromHtml("#DCE3EC");
        private static readonly Color CText = ColorTranslator.FromHtml("#1F2937");
        private static readonly Color CMuted = ColorTranslator.FromHtml("#6B7280");
        private static readonly Color CAccent = ColorTranslator.FromHtml("#2563EB");
        private static readonly Color CAccentDark = ColorTranslator.FromHtml("#1D4ED8");
        private static readonly Color CSelectBg = ColorTranslator.FromHtml("#E8F0FE");
        private static readonly Color CErr = ColorTranslator.FromHtml("#B91C1C");

        private const int DefaultW = 820;
        private const int DefaultH = 520;
        private const int RowH = 44;

        private readonly SnippetStore _store;
        private readonly Func<string, bool> _onInsert;
        public Action OnClosedCallback;                  // 收起之后（把焦点还给 VBE）

        private TextBox _search;
        private SnippetList _list;
        private TextBox _title;
        private TextBox _code;
        private Label _status;
        private Button _btnInsert;

        private List<KeyValuePair<string, string>> _rows = new List<KeyValuePair<string, string>>();
        private List<SnippetHit> _view = new List<SnippetHit>();
        private string _shownTitle;
        private string _shownCode;
        private bool _bulk;                              // 程序性改选中，别当用户点击

        // 这三处字体是我们自己 new 的，必须自己释放 —— Control **不拥有**
        //   Font（HFONT 可以共享），Form.Dispose 不会碰 Form.Font / 子控件的 Font。
        //   不存字段就没法在 Dispose 里释放，每开一次片段窗漏三个 GDI 对象。
        private Font _uiFont;
        private Font _markFont;
        private Font _codeFont;

        public SnippetForm(SnippetStore store, Func<string, bool> onInsert)
        {
            _store = store ?? new SnippetStore();
            _onInsert = onInsert;

            // 标题带上库文件名：一眼就能确认当前在用的是哪个库。
            Text = "代码片段 — " + System.IO.Path.GetFileName(SnippetStore.DbPath());
            try { Icon = IconAssets.Window(); } catch { }
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = CBody;
            _uiFont = new Font("Microsoft YaHei UI", 9f);
            Font = _uiFont;
            AutoScaleMode = AutoScaleMode.Font;
            MinimumSize = new Size(640, 360);
            Size = new Size(DefaultW, DefaultH);
            KeyPreview = true;
            DoubleBuffered = true;

            BuildUi();
            FormClosing += (s, e) =>
            {
                if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Close_(); }
            };
        }

        /// <summary>释放自己 new 的三个 Font（Control 不拥有它们，见字段处说明）。</summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_uiFont != null) { _uiFont.Dispose(); _uiFont = null; }
                if (_markFont != null) { _markFont.Dispose(); _markFont = null; }
                if (_codeFont != null) { _codeFont.Dispose(); _codeFont = null; }
            }
            base.Dispose(disposing);
        }

        // ==================================================================
        // 建界面
        // ==================================================================

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = CBody,
                Margin = new Padding(0),
                Padding = new Padding(0),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            // 工具栏 / 底部固定高度：给 AutoSize 行配 Dock=Fill 的子面板时，
            //   TableLayoutPanel 会算出一个荒唐的高度（实测把搜索框撑到 110px 高）。
            //   固定高度是最省事、也最不会随 DPI 崩掉的做法。
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));        // 工具栏
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));       // 主体
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));        // 状态 + 按钮

            root.Controls.Add(BuildToolbar(), 0, 0);
            root.Controls.Add(BuildMain(), 0, 1);
            root.Controls.Add(BuildFooter(), 0, 2);

            Controls.Add(root);
        }

        /// <summary>
        /// 顶部工具栏：左边"片段库"标签（带强调色竖条）+ 搜索框（占满剩余宽度）。
        /// 刻意不再写"代码片段"四个字 —— 那是窗口标题，重复一遍只是噪音。
        /// 也不在右上角挂"共 N 条 · 命中 M 条"：那是一眼即忘的数字，却会把搜索框
        ///   挤短、还让它的右边缘跟下面两栏对不齐。计数只在底部状态条出现一次。
        /// </summary>
        private Control BuildToolbar()
        {
            var bar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = CBody,
                Padding = new Padding(12, 0, 12, 0),
                Margin = new Padding(0),
            };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            bar.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            var lblMark = new Label
            {
                Text = "片段库",
                AutoSize = true,
                Font = _markFont = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold),
                ForeColor = CText,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(9, 0, 14, 0),
            };
            lblMark.Paint += (s, e) =>
            {
                using (var b = new SolidBrush(CAccent))
                    e.Graphics.FillRectangle(b, new Rectangle(0, 3, 3, Math.Max(6, lblMark.Height - 6)));
            };

            _search = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = CCard,
                ForeColor = CText,
                PlaceholderText = "搜索标题…（多个词用空格分开，↓ 进列表，Enter 插入）",
                // 上下各 15px = 在 56px 的工具栏里垂直居中；右边 0 = 与主体/底栏的
                // 12px 内边距对齐（搜索框右边缘正好压在同一条竖线上）。
                Margin = new Padding(0, 15, 0, 15),
            };
            _search.TextChanged += (s, e) => Refilter(_shownTitle);
            _search.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Down)
                {
                    if (_view.Count == 0) return;
                    if (_list.SelectedIndex < 0) _list.SelectedIndex = 0;
                    _list.Focus();
                    e.Handled = true; e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Enter && !e.Control)
                {
                    DoInsert();
                    e.Handled = true; e.SuppressKeyPress = true;
                }
            };

            bar.Controls.Add(lblMark, 0, 0);
            bar.Controls.Add(_search, 1, 0);
            return bar;
        }

        /// <summary>主体：左列表 + 右编辑区。</summary>
        private Control BuildMain()
        {
            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0),
                Padding = new Padding(12, 0, 12, 0),
                BackColor = CBody,
            };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250f));
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            _list = new SnippetList { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 10, 0) };
            _list.SelectedIndexChanged += (s, e) => OnPick();
            _list.Activate += () => DoInsert();
            _list.TypeToSearch += ch =>
            {
                _search.Focus();
                _search.Text += ch;
                _search.SelectionStart = _search.Text.Length;
            };
            main.Controls.Add(_list, 0, 0);

            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Margin = new Padding(0),
                BackColor = CBody,
            };
            right.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            var lblTitle = new Label
            {
                Text = "标题",
                AutoSize = true,
                ForeColor = CMuted,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 6, 8, 8),
            };
            _title = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = CCard,
                ForeColor = CText,
                Margin = new Padding(0, 2, 0, 8),
            };
            right.Controls.Add(lblTitle, 0, 0);
            right.Controls.Add(_title, 1, 0);

            _code = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                AcceptsTab = true,          // Tab 是缩进，不是切焦点
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = CCard,
                ForeColor = CText,
                Font = _codeFont = new Font("Consolas", 10.5f),
                Margin = new Padding(0),
            };
            _code.KeyDown += CodeKeyDown;
            right.Controls.Add(_code, 0, 1);
            right.SetColumnSpan(_code, 2);

            main.Controls.Add(right, 1, 0);
            return main;
        }

        /// <summary>底部：状态文字在左，按钮在右（"插入"是唯一的主按钮）。</summary>
        private Control BuildFooter()
        {
            // 用表格布局而不是 Panel+Dock：Dock 与 AutoSize 混用时，
            //   "靠右停靠 + 自动宽度"的容器会被 Fill 的兄弟抢成 0 宽 —— 按钮
            //   就是这么整排消失的（截图里只剩状态文字）。
            var foot = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = CBody,
                Padding = new Padding(12, 0, 12, 0),
                Margin = new Padding(0),
            };
            foot.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            foot.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            foot.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            var btnRow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0),
                BackColor = CBody,
            };

            btnRow.Controls.Add(MakeButton("关闭", Close_, false));
            btnRow.Controls.Add(MakeButton("删除", DoDelete, false));
            btnRow.Controls.Add(MakeButton("保存", DoSave, false));
            btnRow.Controls.Add(MakeButton("新建", DoAdd, false));

            _btnInsert = MakeButton("插入到光标处", DoInsert, true);
            _btnInsert.Margin = new Padding(12, 0, 0, 0);
            btnRow.Controls.Add(_btnInsert);

            _status = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = CMuted,
                AutoSize = false,
                Margin = new Padding(0),
                Text = "",
            };

            foot.Controls.Add(_status, 0, 0);
            foot.Controls.Add(btnRow, 1, 0);
            return foot;
        }

        private static Button MakeButton(string text, Action onClick, bool primary)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                BackColor = primary ? CAccent : CCard,
                ForeColor = primary ? Color.White : CText,
                Padding = new Padding(10, 5, 10, 5),
                Margin = new Padding(8, 0, 0, 0),
                Cursor = Cursors.Hand,
            };
            b.FlatAppearance.BorderSize = primary ? 0 : 1;
            b.FlatAppearance.BorderColor = CLine;
            if (primary)
            {
                b.FlatAppearance.MouseOverBackColor = CAccentDark;
                b.FlatAppearance.MouseDownBackColor = CAccentDark;
            }
            b.Click += (s, e) => onClick();
            return b;
        }

        // ==================================================================
        // 显示 / 收起
        // ==================================================================

        /// <summary>弹出并定位。rect = 光标屏幕坐标 (x, 上, 下)；拿不到就传 false。</summary>
        public void OpenAtCaret(bool hasRect, int x, int top, int bottom)
        {
            _bulk = true;
            try { _search.Text = ""; }
            finally { _bulk = false; }

            Reload();
            PlaceWindow(hasRect, x, top, bottom);

            if (!Visible) Show();
            TopMost = true;
            Activate();
            BringToFront();

            _search.Focus();
            _search.SelectAll();
        }

        private void PlaceWindow(bool hasRect, int x, int top, int bottom)
        {
            int w = Math.Max(MinimumSize.Width, Width);
            int h = Math.Max(MinimumSize.Height, Height);
            if (!hasRect)
            {
                Native.POINT p;
                if (Native.GetCursorPos(out p)) { x = p.X; top = p.Y; bottom = p.Y + 16; }
                else { x = 120; top = 120; bottom = 136; }
            }

            Screen scr = Screen.FromPoint(new Point(x, Math.Max(0, bottom)));
            Rectangle wa = scr.WorkingArea;

            int py = bottom + 6;
            if (py + h > wa.Bottom) py = top - h - 6;      // 下方放不下 -> 翻到光标上方
            int px = x;
            if (px + w > wa.Right) px = wa.Right - w;
            if (px < wa.Left) px = wa.Left;
            if (py + h > wa.Bottom) py = wa.Bottom - h;
            if (py < wa.Top) py = wa.Top;

            Bounds = new Rectangle(px, py, w, h);
        }

        /// <summary>收起窗口，并把焦点交回 VBE（否则用户敲键盘会"没反应"）。</summary>
        public void Close_()
        {
            bool wasVisible = Visible;
            Hide();
            if (wasVisible)
            {
                try { OnClosedCallback?.Invoke(); } catch { }
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // 打开期间全局钩子不接管（焦点不在 VBE），所以这两个键必须自己处理
            if (keyData == Keys.Escape || keyData == (Keys.Control | Keys.Q))
            {
                if (IsDirty() && !Confirm("代码还没保存，确定关掉吗？")) return true;
                Close_();
                return true;
            }
            if (keyData == (Keys.Control | Keys.N)) { DoAdd(); return true; }
            if (keyData == (Keys.Control | Keys.S)) { DoSave(); return true; }
            if (keyData == (Keys.Control | Keys.Enter)) { DoInsert(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ==================================================================
        // 列表 / 内容
        // ==================================================================

        /// <summary>从库里重读全部片段，再按当前搜索词过滤。</summary>
        public void Reload(string keep = null)
        {
            try { _rows = _store.All(); }
            catch (Exception e)
            {
                _rows = new List<KeyValuePair<string, string>>();
                Log.Write("读片段库失败：{0}", e.Message);
            }
            Refilter(keep ?? _shownTitle);

            // 库打不开时 All() 只回一张空表，界面上就是"共 0 条"、列表全空 ——
            //   用户会以为片段全丢了。把真正的原因摆到状态条上（覆盖掉条数提示）。
            string err = _store.LastError();
            if (!string.IsNullOrEmpty(err)) SetStatus("片段库打不开：" + err, true);
        }

        private void Refilter(string keep)
        {
            string q = _search == null ? "" : _search.Text;
            _view = SnippetMatch.Filter(_rows, q);

            int target = -1;
            if (!string.IsNullOrEmpty(keep))
            {
                for (int i = 0; i < _view.Count; i++)
                    if (string.Equals(_view[i].Title, keep, StringComparison.Ordinal)) { target = i; break; }
            }
            if (target < 0 && _view.Count > 0) target = 0;

            // 选中动作必须也在 _bulk 里：`SelectedIndex = n` 会触发 OnPick，而 OnPick
            //   在"代码框有未保存改动"时会弹模态确认框 —— 用户每敲一个搜索字符都可能
            //   被弹一次（当前项恰好被过滤掉时）。选中只是刷新列表的结果，不是用户
            //   在切换片段，绝不能走"切换会丢修改吗"那条路修）。
            _bulk = true;
            try
            {
                _list.SetItems(_view);
                if (target >= 0) _list.SelectedIndex = target;
            }
            finally { _bulk = false; }

            if (target >= 0)
            {
                // 文本框里有未保存的改动时**不要**被搜索刷新覆盖掉
                if (!IsDirty()) ShowSnippet(_view[target].Title);
            }
            else if (IsDirty())
            {
                // 一条都没搜到时**绝不能清编辑框**。用户可能正写着一半没
                //   保存的片段，在搜索框里敲词时逐字符都会经过"零命中"这个状态
                //   （打 "msg" 的中间态 "m"、"ms" 常常零命中），否则会顺手把
                //   _title/_code 清空 —— 一整段没保存的代码就这么没了，连确认框
                //   都没有（对比 target>=0 那条，那里是特意用 IsDirty 保护过的）。
                //   列表可以空，编辑器必须留在原地，并说清楚"内容没变、只是没搜到"。
                SetStatus("没有匹配的片段 —— 编辑框里的内容还在（"
                          + (_shownTitle ?? "未保存的新片段") + "），按 Esc 可退出", true);
                return;
            }
            else
            {
                _shownTitle = null;
                _shownCode = null;
                _title.Text = "";
                _code.Text = "";
            }

            SetStatus("共 " + _rows.Count + " 条，命中 " + _view.Count + " 条");
        }

        private void OnPick()
        {
            if (_bulk) return;
            int i = _list.SelectedIndex;
            if (i < 0 || i >= _view.Count) return;
            string title = _view[i].Title;
            if (string.Equals(title, _shownTitle, StringComparison.Ordinal))
            {
                _title.Text = title ?? "";      // 标题框被手改过又点回自己：还原
                return;
            }
            if (IsDirty())
            {
                if (!Confirm("当前文本框里的内容改过但没保存，切换会丢掉这些修改。\n\n要继续切换吗？"))
                {
                    ReselectShown();
                    return;
                }
            }
            ShowSnippet(title);
        }

        private void ReselectShown()
        {
            for (int i = 0; i < _view.Count; i++)
            {
                if (string.Equals(_view[i].Title, _shownTitle, StringComparison.Ordinal))
                {
                    _list.SelectedIndex = i;
                    return;
                }
            }
        }

        /// <summary>文本框内容是否与"当前展示的那条"不一致（纯比较，不靠事件记账）。</summary>
        private bool IsDirty()
        {
            if (_shownTitle == null) return CodeText().Trim().Length > 0;
            return !string.Equals(CodeText(), _shownCode ?? "", StringComparison.Ordinal);
        }

        private string CodeText() => _code == null ? "" : (_code.Text ?? "");

        /// <summary>
        /// 库存内容一律是 LF，而 WinForms 的多行 TextBox 只认 CRLF —— 不换算的话
        /// 整段代码会挤成一行。存 / 插之前再换回 LF，两边字节形态保持一致。
        /// </summary>
        private static string ToCrlf(string s) =>
            (s ?? "").Replace("\r\n", "\n").Replace("\n", "\r\n");

        private static string ToLf(string s) =>
            (s ?? "").Replace("\r\n", "\n").Replace('\r', '\n');

        private void ShowSnippet(string title)
        {
            string code = ToCrlf(_store.Get(title));
            _shownTitle = title;
            _shownCode = code;
            _code.Text = code;
            _title.Text = title ?? "";
            _code.SelectionStart = 0;
            _code.SelectionLength = 0;
        }

        private void SetStatus(string msg, bool err = false)
        {
            if (_status == null) return;
            _status.Text = msg ?? "";
            _status.ForeColor = err ? CErr : CMuted;
        }

        private bool Confirm(string text)
        {
            return MessageBox.Show(this, text, "代码片段", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
        }

        private void CodeKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                DoInsert();
                e.Handled = true; e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.Tab && !e.Control && !e.Alt)
            {
                // 文本框里的 Tab 应该是"缩进"，不是"换焦点"
                int at = _code.SelectionStart;
                _code.SelectedText = "    ";
                _code.SelectionStart = at + 4;
                e.Handled = true; e.SuppressKeyPress = true;
            }
        }

        // ==================================================================
        // 四个动作
        // ==================================================================

        private void DoAdd()
        {
            string title = (_title.Text ?? "").Trim();
            string code = CodeText();
            if (title.Length == 0)
            {
                SetStatus("先填标题，再点「新建」", true);
                _title.Focus();
                return;
            }
            if (code.Trim().Length == 0)
            {
                SetStatus("代码是空的，没法存", true);
                return;
            }
            if (_store.Add(title, ToLf(code)))
            {
                Reload(title);
                ShowSnippet(title);
                SetStatus("已新增「" + title + "」");
            }
            else if (_store.Exists(title))
            {
                SetStatus("标题「" + title + "」已存在 —— 想改它请用「保存」", true);
            }
            else
            {
                // 数据库写失败（文件被占用 / 只读目录 / 磁盘满）不该说成"标题重复"
                SetStatus("新增失败：" + _store.LastError(), true);
            }
        }

        private void DoSave()
        {
            string cur = _shownTitle;
            if (string.IsNullOrEmpty(cur))
            {
                SetStatus("先在左边选一条要修改的片段", true);
                return;
            }
            string title = (_title.Text ?? "").Trim();
            string code = CodeText();
            if (title.Length == 0)
            {
                SetStatus("标题不能为空", true);
                return;
            }
            if (!string.Equals(title, cur, StringComparison.Ordinal))
            {
                if (_store.Exists(title))
                {
                    SetStatus("标题「" + title + "」已被别的片段占用", true);
                    return;
                }
                if (!_store.Rename(cur, title))
                {
                    SetStatus("改标题失败：" + _store.LastError(), true);
                    return;
                }
                // 标题已经改掉了：这一步失败也必须把界面刷成"新标题 + 库里的代码"，
                //   否则用户看到的是"标题变了、代码还是旧的"，还以为保存成功。
                if (!_store.Update(title, ToLf(code)))
                {
                    Reload(title);
                    ShowSnippet(title);
                    SetStatus("标题已改，但代码没存进去：" + _store.LastError(), true);
                    return;
                }
                Reload(title);
                ShowSnippet(title);
                SetStatus("已保存（标题改为「" + title + "」）");
                return;
            }
            if (_store.Update(title, ToLf(code)))
            {
                Reload(title);
                ShowSnippet(title);
                SetStatus("已保存「" + title + "」");
            }
            else
            {
                SetStatus("保存失败（这条片段可能已被删掉）", true);
            }
        }

        private void DoDelete()
        {
            string cur = _shownTitle;
            if (string.IsNullOrEmpty(cur))
            {
                SetStatus("先在左边选一条要删除的片段", true);
                return;
            }
            if (!Confirm("确定删除片段「" + cur + "」？\n\n此操作不可撤销。")) return;

            if (_store.Delete(cur))
            {
                _shownTitle = null;
                _shownCode = null;
                _code.Text = "";
                _title.Text = "";
                Reload(null);
                SetStatus("已删除「" + cur + "」");
            }
            else
            {
                SetStatus("删除失败", true);
            }
        }

        private void DoInsert()
        {
            string code = CodeText();
            if (code.Trim().Length == 0)
            {
                SetStatus("没有可插入的代码", true);
                return;
            }
            bool ok = false;
            try { ok = _onInsert != null && _onInsert(code); }
            catch (Exception e) { Log.Write("插入片段异常：{0}", e.Message); }
            if (!ok)
            {
                SetStatus("插入失败 —— 光标是不是已经离开代码窗格了？", true);
                return;
            }
            Close_();                       // 收起（OnClosed 里把焦点交回 VBE）
        }

        // ==================================================================
        // 给 --uitest 用：把窗子离屏画成一张图，用来核对布局配色
        // ==================================================================
        public void PreviewOnly(List<KeyValuePair<string, string>> rows)
        {
            _rows = rows ?? new List<KeyValuePair<string, string>>();
            _search.Text = "";
            Refilter(null);
        }

        // ==================================================================
        // 自绘列表：两行一条（标题 + 代码首行），搜索命中逐字符高亮
        // ==================================================================
        private sealed class SnippetList : Control
        {
            private readonly List<SnippetHit> _items = new List<SnippetHit>();
            private int _sel = -1;
            private int _top;
            private int _hover = -1;

            public event EventHandler SelectedIndexChanged;
            public event Action Activate;
            public event Action<char> TypeToSearch;

            public SnippetList()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
                BackColor = CCard;
                ForeColor = CText;
                TabStop = true;
            }

            // 这个特性不是装饰：自绘控件上的属性会被 WinForms 的**设计器序列化**
            //   分析器检查（WFO1000），缺了它整次 publish 直接失败（error，不是 warning），
            //   而失败日志里那条错又很容易被 --uitest 的"看起来正常"掩盖过去。
            [System.ComponentModel.DesignerSerializationVisibility(
                System.ComponentModel.DesignerSerializationVisibility.Hidden)]
            public int SelectedIndex
            {
                get => _sel;
                set
                {
                    int v = _items.Count == 0 ? -1 : Math.Max(0, Math.Min(_items.Count - 1, value));
                    if (v == _sel) return;
                    _sel = v;
                    EnsureVisible();
                    Invalidate();
                    SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
                }
            }

            public void SetItems(List<SnippetHit> items)
            {
                _items.Clear();
                if (items != null) _items.AddRange(items);
                _sel = -1;
                _top = 0;
                Invalidate();
            }

            private int VisibleRows => Math.Max(1, (Height - 2) / RowH);

            private void EnsureVisible()
            {
                if (_sel < 0) return;
                if (_sel < _top) _top = _sel;
                else if (_sel >= _top + VisibleRows) _top = _sel - VisibleRows + 1;
                int maxTop = Math.Max(0, _items.Count - VisibleRows);
                if (_top > maxTop) _top = maxTop;
                if (_top < 0) _top = 0;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var bg = new SolidBrush(CCard)) g.FillRectangle(bg, ClientRectangle);

                if (_items.Count == 0)
                {
                    using (var b = new SolidBrush(CMuted))
                    using (var f = new Font("Microsoft YaHei UI", 9f))
                    {
                        string t = "还没有片段 —— 在右边写好代码，点「新建」";
                        var sz = g.MeasureString(t, f);
                        g.DrawString(t, f, b, (Width - sz.Width) / 2, (Height - sz.Height) / 2);
                    }
                    using (var pen = new Pen(CLine)) g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                    return;
                }

                int rows = VisibleRows;
                int last = Math.Min(_items.Count, _top + rows);
                using (var titleFont = new Font("Microsoft YaHei UI", 9.5f))
                using (var titleBold = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold))
                using (var subFont = new Font("Consolas", 8.5f))
                using (var textBrush = new SolidBrush(CText))
                using (var mutBrush = new SolidBrush(CMuted))
                using (var hitBrush = new SolidBrush(CAccentDark))
                using (var selBrush = new SolidBrush(CSelectBg))
                using (var barBrush = new SolidBrush(CAccent))
                using (var hovBrush = new SolidBrush(ColorTranslator.FromHtml("#F3F6FB")))
                using (var linePen = new Pen(ColorTranslator.FromHtml("#EEF1F5")))
                {
                    for (int i = _top; i < last; i++)
                    {
                        int y = (i - _top) * RowH;
                        bool sel = i == _sel;

                        if (sel) g.FillRectangle(selBrush, new Rectangle(0, y, Width, RowH));
                        else if (i == _hover) g.FillRectangle(hovBrush, new Rectangle(0, y, Width, RowH));
                        if (sel) g.FillRectangle(barBrush, new Rectangle(0, y, 3, RowH));

                        var it = _items[i];
                        string title = it.Title ?? "";
                        float ty = y + 6f;

                        // ---- 标题：逐字符高亮（命中的字符单独用强调色画） ----
                        var hits = it.Hits;
                        if (hits == null || hits.Count == 0)
                        {
                            g.DrawString(title, sel ? titleBold : titleFont, textBrush, 12f, ty);
                        }
                        else
                        {
                            float x = 12f;
                            for (int k = 0; k < title.Length; k++)
                            {
                                string ch = title.Substring(k, 1);
                                bool on = hits.Contains(k);
                                var f = on ? titleBold : titleFont;
                                var br = on ? hitBrush : textBrush;
                                g.DrawString(ch, f, br, x, ty);
                                x += g.MeasureString(ch, f).Width - 1.2f;   // 逐字拼接收紧一点
                            }
                        }

                        // ---- 副行：代码第一行，等宽、灰色 ----
                        string first = FirstLine(it.Code);
                        if (first.Length > 0)
                        {
                            var sz = g.MeasureString(first, subFont);
                            if (sz.Width > Width - 20)
                            {
                                // 超宽就按比例截断加省略号（不做逐字裁剪，够用且快）
                                int keep = Math.Max(1, (int)(first.Length * (Width - 30) / sz.Width));
                                first = first.Substring(0, Math.Min(keep, first.Length)) + "…";
                            }
                            g.DrawString(first, subFont, mutBrush, 12f, y + 24f);
                        }

                        if (i < last - 1)
                            g.DrawLine(linePen, 12, y + RowH - 1, Width - 4, y + RowH - 1);
                    }

                    // 右侧滚动条（只在需要时出现）
                    if (_items.Count > rows)
                    {
                        int trackH = Height - 4;
                        int thumbH = Math.Max(24, trackH * rows / _items.Count);
                        int maxTop = Math.Max(1, _items.Count - rows);
                        int thumbY = 2 + (trackH - thumbH) * _top / maxTop;
                        using (var tb = new SolidBrush(ColorTranslator.FromHtml("#D8DEE8")))
                            g.FillRectangle(tb, new Rectangle(Width - 6, thumbY, 4, thumbH));
                    }
                }

                using (var pen = new Pen(CLine)) g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            }

            private static string FirstLine(string code)
            {
                if (string.IsNullOrEmpty(code)) return "";
                int i = code.IndexOfAny(new[] { '\r', '\n' });
                string s = i < 0 ? code : code.Substring(0, i);
                return s.Trim();
            }

            private int RowAt(int y) => _top + y / RowH;

            protected override void OnMouseMove(MouseEventArgs e)
            {
                base.OnMouseMove(e);
                int r = RowAt(e.Y);
                if (r >= _items.Count) r = -1;
                if (r != _hover) { _hover = r; Invalidate(); }
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                if (_hover != -1) { _hover = -1; Invalidate(); }
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                base.OnMouseDown(e);
                Focus();
                int r = RowAt(e.Y);
                if (r >= 0 && r < _items.Count)
                {
                    SelectedIndex = r;
                    if (e.Clicks >= 2) Activate?.Invoke();
                }
            }

            protected override void OnMouseWheel(MouseEventArgs e)
            {
                base.OnMouseWheel(e);
                int maxTop = Math.Max(0, _items.Count - VisibleRows);
                int nt = Math.Max(0, Math.Min(maxTop, _top - Math.Sign(e.Delta) * 3));
                if (nt != _top) { _top = nt; Invalidate(); }
            }

            protected override bool IsInputKey(Keys keyData)
            {
                // 让上下键 / 回车进到 OnKeyDown，而不是被父容器当成"切换焦点"
                switch (keyData)
                {
                    case Keys.Up:
                    case Keys.Down:
                    case Keys.PageUp:
                    case Keys.PageDown:
                    case Keys.Home:
                    case Keys.End:
                    case Keys.Enter:
                        return true;
                }
                return base.IsInputKey(keyData);
            }

            protected override void OnKeyDown(KeyEventArgs e)
            {
                base.OnKeyDown(e);
                switch (e.KeyCode)
                {
                    case Keys.Up: SelectedIndex = _sel - 1; e.Handled = true; break;
                    case Keys.Down: SelectedIndex = _sel + 1; e.Handled = true; break;
                    case Keys.PageUp: SelectedIndex = _sel - VisibleRows; e.Handled = true; break;
                    case Keys.PageDown: SelectedIndex = _sel + VisibleRows; e.Handled = true; break;
                    case Keys.Home: SelectedIndex = 0; e.Handled = true; break;
                    case Keys.End: SelectedIndex = _items.Count - 1; e.Handled = true; break;
                    case Keys.Enter:
                        if (!e.Shift && !e.Control) { Activate?.Invoke(); e.Handled = true; e.SuppressKeyPress = true; }
                        break;
                }
            }

            protected override void OnKeyPress(KeyPressEventArgs e)
            {
                base.OnKeyPress(e);
                // 在列表里直接打字 -> 把这一下送进搜索框（别让列表自己跳字）
                if (!char.IsControl(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar))
                {
                    TypeToSearch?.Invoke(e.KeyChar);
                    e.Handled = true;
                }
            }
        }
    }
}
