using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Ribbon.Vsto
{
    internal static class RibbonUiThread
    {
        public static bool Run(Control owner, Action action)
        {
            if (owner == null || action == null || owner.IsDisposed || owner.Disposing || !owner.IsHandleCreated) return false;
            try
            {
                if (!owner.InvokeRequired)
                {
                    if (owner.IsDisposed || owner.Disposing || !owner.IsHandleCreated) return false;
                    action();
                    return true;
                }

                owner.Invoke(new Action(() =>
                {
                    if (!owner.IsDisposed && !owner.Disposing && owner.IsHandleCreated) action();
                }));
                return true;
            }
            catch (ObjectDisposedException)
            {
                return false;
            }
            catch (InvalidOperationException) when (owner.IsDisposed || owner.Disposing || !owner.IsHandleCreated)
            {
                return false;
            }
        }

        public static void Post(Control owner, Action action)
        {
            if (owner == null || action == null || owner.IsDisposed || owner.Disposing || !owner.IsHandleCreated) return;
            try
            {
                if (!owner.InvokeRequired)
                {
                    if (!owner.IsDisposed && !owner.Disposing && owner.IsHandleCreated) action();
                    return;
                }

                owner.BeginInvoke(new Action(() =>
                {
                    if (!owner.IsDisposed && !owner.Disposing && owner.IsHandleCreated) action();
                }));
            }
            catch (ObjectDisposedException)
            {
            }
            catch (InvalidOperationException) when (owner.IsDisposed || owner.Disposing || !owner.IsHandleCreated)
            {
            }
        }
    }

    internal sealed class RibbonPalette
    {
        private RibbonPalette(bool dark, Color background, Color surface, Color surfaceRaised, Color border,
            Color text, Color mutedText, Color accent, Color accentHover, Color success, Color danger)
        {
            IsDark = dark;
            Background = background;
            Surface = surface;
            SurfaceRaised = surfaceRaised;
            Border = border;
            Text = text;
            MutedText = mutedText;
            Accent = accent;
            AccentHover = accentHover;
            Success = success;
            Danger = danger;
        }

        public bool IsDark { get; }
        public Color Background { get; }
        public Color Surface { get; }
        public Color SurfaceRaised { get; }
        public Color Border { get; }
        public Color Text { get; }
        public Color MutedText { get; }
        public Color Accent { get; }
        public Color AccentHover { get; }
        public Color Success { get; }
        public Color Danger { get; }

        public static RibbonPalette Detect()
        {
            if (SystemInformation.HighContrast)
            {
                return new RibbonPalette(
                    SystemColors.Window.GetBrightness() < 0.5f,
                    SystemColors.Window,
                    SystemColors.Window,
                    SystemColors.Control,
                    SystemColors.WindowFrame,
                    SystemColors.WindowText,
                    SystemColors.GrayText,
                    SystemColors.Highlight,
                    SystemColors.HotTrack,
                    Color.LimeGreen,
                    Color.OrangeRed);
            }

            var dark = IsDarkAppMode();
            return dark
                ? new RibbonPalette(
                    true,
                    Color.FromArgb(24, 26, 31),
                    Color.FromArgb(31, 34, 40),
                    Color.FromArgb(40, 44, 52),
                    Color.FromArgb(58, 63, 73),
                    Color.FromArgb(244, 246, 249),
                    Color.FromArgb(166, 173, 185),
                    Color.FromArgb(104, 119, 245),
                    Color.FromArgb(124, 137, 250),
                    Color.FromArgb(70, 201, 142),
                    Color.FromArgb(236, 107, 117))
                : new RibbonPalette(
                    false,
                    Color.FromArgb(246, 247, 250),
                    Color.White,
                    Color.FromArgb(238, 241, 246),
                    Color.FromArgb(216, 220, 228),
                    Color.FromArgb(31, 35, 43),
                    Color.FromArgb(103, 111, 126),
                    Color.FromArgb(79, 99, 232),
                    Color.FromArgb(65, 84, 214),
                    Color.FromArgb(37, 155, 105),
                    Color.FromArgb(204, 72, 83));
        }

        private static bool IsDarkAppMode()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    var value = key?.GetValue("AppsUseLightTheme");
                    return value != null && Convert.ToInt32(value) == 0;
                }
            }
            catch
            {
                return false;
            }
        }
    }

    internal enum RibbonButtonKind
    {
        Primary,
        Secondary,
        Ghost,
        Danger
    }

    internal enum RibbonGlyph
    {
        None,
        Agents,
        Send,
        Stop,
        Refresh,
        Download,
        Remove
    }

    internal sealed class RibbonButton : Button
    {
        private readonly RibbonPalette _palette;
        private bool _hovered;
        private bool _pressed;
        private RibbonButtonKind _kind;

        public RibbonButton(RibbonPalette palette, RibbonButtonKind kind = RibbonButtonKind.Secondary)
        {
            _palette = palette;
            _kind = kind;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            UseVisualStyleBackColor = false;
            UseCompatibleTextRendering = false;
            AutoSize = false;
            Height = 34;
            MinimumSize = new Size(72, 34);
            Cursor = Cursors.Hand;
            TabStop = true;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        public RibbonButtonKind Kind
        {
            get { return _kind; }
            set { _kind = value; Invalidate(); }
        }

        public RibbonGlyph Glyph { get; set; }

        protected override void OnMouseEnter(EventArgs e) { _hovered = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hovered = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs mevent) { _pressed = true; Invalidate(); base.OnMouseDown(mevent); }
        protected override void OnMouseUp(MouseEventArgs mevent) { _pressed = false; Invalidate(); base.OnMouseUp(mevent); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var background = new SolidBrush(Parent?.BackColor ?? _palette.Surface))
            {
                e.Graphics.FillRectangle(background, ClientRectangle);
            }
            var bounds = new Rectangle(0, 0, Width - 1, Height - 1);
            var colors = GetColors();
            using (var path = RibbonDrawing.RoundRectangle(bounds, 7))
            using (var brush = new SolidBrush(colors.Item1))
            using (var pen = new Pen(colors.Item2))
            {
                e.Graphics.FillPath(brush, path);
                e.Graphics.DrawPath(pen, path);
            }

            var textColor = Enabled ? colors.Item3 : RibbonDrawing.Blend(_palette.MutedText, colors.Item1, 0.34f);
            var textSize = TextRenderer.MeasureText(Text ?? string.Empty, Font, Size.Empty, TextFormatFlags.NoPadding);
            var hasGlyph = Glyph != RibbonGlyph.None;
            var glyphWidth = hasGlyph ? 16 : 0;
            var gap = hasGlyph && !string.IsNullOrEmpty(Text) ? 7 : 0;
            var contentWidth = glyphWidth + gap + textSize.Width;
            var left = Math.Max(7, (Width - contentWidth) / 2);
            if (hasGlyph)
            {
                DrawGlyph(e.Graphics, new Rectangle(left, (Height - 16) / 2, 16, 16), textColor);
                left += glyphWidth + gap;
            }
            TextRenderer.DrawText(e.Graphics, Text ?? string.Empty, Font,
                new Rectangle(left, 0, Math.Max(0, Width - left - 6), Height), textColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);

            if (Focused && ShowFocusCues)
            {
                ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -4, -4), textColor, colors.Item1);
            }
        }

        private Tuple<Color, Color, Color> GetColors()
        {
            Color fill;
            Color border;
            Color text;
            switch (_kind)
            {
                case RibbonButtonKind.Primary:
                    fill = _hovered ? _palette.AccentHover : _palette.Accent;
                    border = fill;
                    text = Color.White;
                    break;
                case RibbonButtonKind.Danger:
                    fill = _hovered ? Color.FromArgb(38, _palette.Danger) : Color.FromArgb(22, _palette.Danger);
                    border = Color.FromArgb(150, _palette.Danger);
                    text = _palette.Danger;
                    break;
                case RibbonButtonKind.Ghost:
                    fill = _hovered ? _palette.SurfaceRaised : _palette.Surface;
                    border = _hovered ? _palette.Border : _palette.Surface;
                    text = _palette.Text;
                    break;
                default:
                    fill = _hovered ? _palette.SurfaceRaised : _palette.Surface;
                    border = _palette.Border;
                    text = _palette.Text;
                    break;
            }
            if (_pressed) fill = RibbonDrawing.Blend(fill, Color.Black, _palette.IsDark ? 0.16f : 0.08f);
            if (!Enabled)
            {
                fill = RibbonDrawing.Blend(fill, _palette.Surface, 0.68f);
                border = RibbonDrawing.Blend(border, _palette.Surface, 0.58f);
                text = RibbonDrawing.Blend(text, _palette.Surface, 0.5f);
            }
            return Tuple.Create(fill, border, text);
        }

        private void DrawGlyph(Graphics graphics, Rectangle bounds, Color color)
        {
            using (var pen = new Pen(color, 1.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                switch (Glyph)
                {
                    case RibbonGlyph.Agents:
                        graphics.DrawEllipse(pen, bounds.Left + 5, bounds.Top + 2, 6, 6);
                        graphics.DrawArc(pen, bounds.Left + 3, bounds.Top + 9, 10, 7, 190, 160);
                        break;
                    case RibbonGlyph.Send:
                        graphics.DrawLine(pen, bounds.Left + 2, bounds.Top + 8, bounds.Right - 3, bounds.Top + 8);
                        graphics.DrawLine(pen, bounds.Right - 7, bounds.Top + 4, bounds.Right - 3, bounds.Top + 8);
                        graphics.DrawLine(pen, bounds.Right - 7, bounds.Top + 12, bounds.Right - 3, bounds.Top + 8);
                        break;
                    case RibbonGlyph.Stop:
                        graphics.DrawRectangle(pen, bounds.Left + 4, bounds.Top + 4, 8, 8);
                        break;
                    case RibbonGlyph.Refresh:
                        graphics.DrawArc(pen, bounds.Left + 3, bounds.Top + 3, 10, 10, 35, 285);
                        graphics.DrawLine(pen, bounds.Left + 11, bounds.Top + 2, bounds.Right - 2, bounds.Top + 3);
                        graphics.DrawLine(pen, bounds.Right - 2, bounds.Top + 3, bounds.Right - 4, bounds.Top + 6);
                        break;
                    case RibbonGlyph.Download:
                        graphics.DrawLine(pen, bounds.Left + 8, bounds.Top + 2, bounds.Left + 8, bounds.Top + 10);
                        graphics.DrawLine(pen, bounds.Left + 5, bounds.Top + 7, bounds.Left + 8, bounds.Top + 10);
                        graphics.DrawLine(pen, bounds.Left + 11, bounds.Top + 7, bounds.Left + 8, bounds.Top + 10);
                        graphics.DrawLine(pen, bounds.Left + 3, bounds.Top + 13, bounds.Right - 3, bounds.Top + 13);
                        break;
                    case RibbonGlyph.Remove:
                        graphics.DrawLine(pen, bounds.Left + 4, bounds.Top + 8, bounds.Right - 4, bounds.Top + 8);
                        break;
                }
            }
        }
    }

    internal sealed class RibbonSurface : Panel
    {
        private readonly RibbonPalette _palette;
        private bool _emphasizeBorder;
        private bool _useRaisedBackground;

        public RibbonSurface(RibbonPalette palette)
        {
            _palette = palette;
            BackColor = palette.Surface;
            Padding = new Padding(1);
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        public int CornerRadius { get; set; } = 9;

        public bool EmphasizeBorder
        {
            get { return _emphasizeBorder; }
            set
            {
                if (_emphasizeBorder == value) return;
                _emphasizeBorder = value;
                Invalidate();
            }
        }

        public bool UseRaisedBackground
        {
            get { return _useRaisedBackground; }
            set
            {
                if (_useRaisedBackground == value) return;
                _useRaisedBackground = value;
                BackColor = value ? _palette.SurfaceRaised : _palette.Surface;
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var bounds = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = RibbonDrawing.RoundRectangle(bounds, CornerRadius))
            using (var fill = new SolidBrush(_useRaisedBackground ? _palette.SurfaceRaised : _palette.Surface))
            using (var border = new Pen(_emphasizeBorder ? _palette.Accent : _palette.Border, _emphasizeBorder ? 1.4f : 1f))
            {
                e.Graphics.FillPath(fill, path);
                e.Graphics.DrawPath(border, path);
            }
            base.OnPaint(e);
        }
    }

    internal sealed class RibbonLayoutPanel : TableLayoutPanel
    {
        public RibbonLayoutPanel()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }
    }

    internal sealed class RibbonComboBox : ComboBox
    {
        private const int WmPaint = 0x000F;
        private const int WmNcPaint = 0x0085;
        private const int WmPrint = 0x0317;
        private const int WmPrintClient = 0x0318;
        private readonly RibbonPalette _palette;

        public RibbonComboBox(RibbonPalette palette)
        {
            _palette = palette;
            FlatStyle = FlatStyle.Flat;
            BackColor = palette.SurfaceRaised;
            ForeColor = palette.Text;
        }

        public string PlaceholderText { get; set; }

        protected override void OnSelectedIndexChanged(EventArgs e)
        {
            base.OnSelectedIndexChanged(e);
            Invalidate();
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            Invalidate();
        }

        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            Invalidate();
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            Invalidate();
        }

        protected override void WndProc(ref Message message)
        {
            base.WndProc(ref message);
            if (message.Msg != WmPaint && message.Msg != WmNcPaint && message.Msg != WmPrint && message.Msg != WmPrintClient) return;
            try
            {
                using (var graphics = (message.Msg == WmPrint || message.Msg == WmPrintClient) && message.WParam != IntPtr.Zero
                    ? Graphics.FromHdc(message.WParam)
                    : Graphics.FromHwnd(Handle))
                using (var background = new SolidBrush(_palette.SurfaceRaised))
                using (var border = new Pen(Focused ? _palette.Accent : _palette.Border, Focused ? 1.4f : 1f))
                using (var arrow = new SolidBrush(Enabled ? _palette.MutedText : RibbonDrawing.Blend(_palette.MutedText, _palette.SurfaceRaised, 0.45f)))
                {
                    var buttonWidth = Math.Max(22, Height);
                    graphics.FillRectangle(background, ClientRectangle);
                    graphics.DrawRectangle(border, 0, 0, Width - 1, Height - 1);
                    var text = GetItemText(SelectedItem);
                    var placeholder = string.IsNullOrWhiteSpace(text);
                    if (placeholder) text = PlaceholderText ?? string.Empty;
                    var textColor = placeholder || !Enabled
                        ? RibbonDrawing.Blend(_palette.MutedText, _palette.SurfaceRaised, !Enabled ? 0.4f : 0f)
                        : _palette.Text;
                    TextRenderer.DrawText(graphics, text ?? string.Empty, Font,
                        new Rectangle(9, 1, Math.Max(0, Width - buttonWidth - 12), Height - 2), textColor,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
                    var centerX = Width - buttonWidth / 2;
                    var centerY = Height / 2;
                    graphics.FillPolygon(arrow, new[]
                    {
                        new Point(centerX - 4, centerY - 2),
                        new Point(centerX + 4, centerY - 2),
                        new Point(centerX, centerY + 3)
                    });
                }
            }
            catch
            {
            }
        }
    }

    internal sealed class RibbonDropArrow : Control
    {
        private readonly RibbonPalette _palette;
        private readonly ComboBox _comboBox;

        public RibbonDropArrow(RibbonPalette palette, ComboBox comboBox)
        {
            _palette = palette;
            _comboBox = comboBox;
            Width = 30;
            Dock = DockStyle.Right;
            Cursor = Cursors.Hand;
            BackColor = palette.SurfaceRaised;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            Click += (sender, args) => { if (_comboBox.Enabled) _comboBox.DroppedDown = true; };
            _comboBox.EnabledChanged += ComboBoxOnVisualStateChanged;
            _comboBox.GotFocus += ComboBoxOnVisualStateChanged;
            _comboBox.LostFocus += ComboBoxOnVisualStateChanged;
        }

        private void ComboBoxOnVisualStateChanged(object sender, EventArgs e)
        {
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var background = new SolidBrush(_palette.SurfaceRaised)) e.Graphics.FillRectangle(background, ClientRectangle);
            using (var arrow = new SolidBrush(_comboBox.Enabled ? _palette.MutedText : RibbonDrawing.Blend(_palette.MutedText, _palette.SurfaceRaised, 0.35f)))
            {
                var centerX = Width / 2;
                var centerY = Height / 2;
                e.Graphics.FillPolygon(arrow, new[]
                {
                    new Point(centerX - 4, centerY - 2),
                    new Point(centerX + 4, centerY - 2),
                    new Point(centerX, centerY + 3)
                });
            }
            using (var border = new Pen(_comboBox.Focused ? _palette.Accent : _palette.Border)) e.Graphics.DrawLine(border, 0, 0, 0, Height);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _comboBox.EnabledChanged -= ComboBoxOnVisualStateChanged;
                _comboBox.GotFocus -= ComboBoxOnVisualStateChanged;
                _comboBox.LostFocus -= ComboBoxOnVisualStateChanged;
            }
            base.Dispose(disposing);
        }
    }

    internal sealed class RibbonBrandMark : Control
    {
        private readonly RibbonPalette _palette;
        private readonly Font _markFont;
        private readonly string _mark;
        private readonly Color _brandColor;

        public RibbonBrandMark(RibbonPalette palette)
            : this(palette, "R", palette.Accent)
        {
        }

        public RibbonBrandMark(RibbonPalette palette, string mark, Color brandColor)
        {
            _palette = palette;
            _mark = string.IsNullOrWhiteSpace(mark) ? "R" : mark.Substring(0, 1).ToUpperInvariant();
            _brandColor = brandColor;
            _markFont = new Font("Segoe UI", 11f, FontStyle.Bold, GraphicsUnit.Point);
            Size = new Size(32, 32);
            MinimumSize = Size;
            MaximumSize = Size;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var highlight = RibbonDrawing.Blend(_brandColor, Color.White, _palette.IsDark ? 0.22f : 0.12f);
            using (var brush = new LinearGradientBrush(ClientRectangle, _brandColor, highlight, 45f))
            {
                e.Graphics.FillEllipse(brush, 0, 0, Width - 1, Height - 1);
            }
            using (var ring = new Pen(Color.FromArgb(_palette.IsDark ? 70 : 45, Color.White)))
            {
                e.Graphics.DrawEllipse(ring, 1, 1, Width - 3, Height - 3);
            }
            TextRenderer.DrawText(e.Graphics, _mark, _markFont, ClientRectangle,
                Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _markFont.Dispose();
            base.Dispose(disposing);
        }
    }

    internal sealed class RibbonStatusDot : Control
    {
        public RibbonStatusDot()
        {
            Size = new Size(12, 12);
            MinimumSize = Size;
            MaximumSize = Size;
            SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        public Color DotColor { get; set; } = Color.Gray;

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var brush = new SolidBrush(DotColor)) e.Graphics.FillEllipse(brush, 2, 2, 8, 8);
        }
    }

    internal static class RibbonWindowChrome
    {
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        public static void Apply(Form form, RibbonPalette palette)
        {
            if (!palette.IsDark || form == null || !form.IsHandleCreated) return;
            try
            {
                var enabled = 1;
                if (DwmSetWindowAttribute(form.Handle, 20, ref enabled, sizeof(int)) != 0)
                {
                    DwmSetWindowAttribute(form.Handle, 19, ref enabled, sizeof(int));
                }
            }
            catch
            {
            }
        }
    }

    internal static class RibbonNativeTheme
    {
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int SetWindowTheme(IntPtr handle, string subAppName, string subIdList);

        public static void ApplyDarkScrollBars(Control control, RibbonPalette palette)
        {
            if (control == null || palette == null || !palette.IsDark || SystemInformation.HighContrast) return;
            control.HandleCreated += (sender, args) => ApplyDarkExplorerTheme(control);
            if (control.IsHandleCreated) ApplyDarkExplorerTheme(control);
        }

        private static void ApplyDarkExplorerTheme(Control control)
        {
            if (control == null || control.IsDisposed || !control.IsHandleCreated) return;
            try
            {
                SetWindowTheme(control.Handle, "DarkMode_Explorer", null);
                control.Invalidate(true);
            }
            catch (DllNotFoundException)
            {
            }
            catch (EntryPointNotFoundException)
            {
            }
        }
    }

    internal static class RibbonCue
    {
        private const int EmSetCueBanner = 0x1501;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr showWhenFocused, string text);

        public static void Set(TextBox textBox, string text)
        {
            if (textBox == null) return;
            EventHandler apply = null;
            apply = (sender, args) =>
            {
                try { SendMessage(textBox.Handle, EmSetCueBanner, new IntPtr(1), text); } catch { }
                textBox.HandleCreated -= apply;
            };
            if (textBox.IsHandleCreated) apply(textBox, EventArgs.Empty);
            else textBox.HandleCreated += apply;
        }
    }

    internal static class RibbonDrawing
    {
        public static GraphicsPath RoundRectangle(Rectangle bounds, int radius)
        {
            var diameter = Math.Max(2, radius * 2);
            var path = new GraphicsPath();
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }

        public static Color Blend(Color first, Color second, float amount)
        {
            amount = Math.Max(0f, Math.Min(1f, amount));
            return Color.FromArgb(
                (int)(first.A + (second.A - first.A) * amount),
                (int)(first.R + (second.R - first.R) * amount),
                (int)(first.G + (second.G - first.G) * amount),
                (int)(first.B + (second.B - first.B) * amount));
        }
    }
}
