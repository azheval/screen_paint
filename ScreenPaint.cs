
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// ----------------------------------------------------------------------------

namespace ScreenPaint
{
    // ------------------------------------------------------------------ Win32

    internal static class Native
    {
        public const int GWL_EXSTYLE = -20;
        public const int WS_EX_LAYERED = 0x00080000;
        public const int WS_EX_TRANSPARENT = 0x00000020;
        public const int WS_EX_TOOLWINDOW = 0x00000080;
        public const int WS_EX_NOACTIVATE = 0x08000000;

        public const int ULW_ALPHA = 2;
        public const byte AC_SRC_OVER = 0;
        public const byte AC_SRC_ALPHA = 1;

        public const int SM_XVIRTUALSCREEN = 76;
        public const int SM_YVIRTUALSCREEN = 77;
        public const int SM_CXVIRTUALSCREEN = 78;
        public const int SM_CYVIRTUALSCREEN = 79;
        public const int SM_CXSCREEN = 0;
        public const int SM_CYSCREEN = 1;
        public const int LOGPIXELSX = 88;

        public const int RGN_DIFF = 4;
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        public const uint SWP_NOSIZE = 0x0001, SWP_NOMOVE = 0x0002, SWP_NOACTIVATE = 0x0010;

        public const uint MOD_ALT = 0x0001, MOD_CONTROL = 0x0002, MOD_NOREPEAT = 0x4000;
        public const int WM_HOTKEY = 0x0312;

        public const int WM_DPICHANGED = 0x02E0;
        public const int WM_DISPLAYCHANGE = 0x007E;

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X, Y;
            public POINT(int x, int y) { X = x; Y = y; }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SIZE
        {
            public int cx, cy;
            public SIZE(int w, int h) { cx = w; cy = h; }
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BLENDFUNCTION
        {
            public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFOHEADER
        {
            public int biSize, biWidth, biHeight;
            public short biPlanes, biBitCount;
            public int biCompression, biSizeImage, biXPelsPerMeter, biYPelsPerMeter, biClrUsed, biClrImportant;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFO
        {
            public BITMAPINFOHEADER bmiHeader;
            public int bmiColors;
        }

        [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("user32.dll")] public static extern int GetSystemMetrics(int index);
        [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
        [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr value);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref POINT pptDst,
            ref SIZE psize, IntPtr hdcSrc, ref POINT pptSrc, int crKey, ref BLENDFUNCTION pblend, int dwFlags);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

        [DllImport("dwmapi.dll")]
        public static extern int DwmFlush();

        [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint modifiers, uint vk);
        [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vKey);
        public const int VK_SHIFT = 0x10;

        [DllImport("gdi32.dll")] public static extern IntPtr CreateDIBSection(IntPtr hdc, ref BITMAPINFO pbmi,
            uint usage, out IntPtr ppvBits, IntPtr hSection, uint offset);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleDC(IntPtr hdc);
        [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hObject);
        [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr hdc);
        [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr hObject);
        [DllImport("gdi32.dll")] public static extern int GetDeviceCaps(IntPtr hdc, int index);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateRectRgn(int l, int t, int r, int b);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateRoundRectRgn(int l, int t, int r, int b, int w, int h);
        [DllImport("gdi32.dll")] public static extern int CombineRgn(IntPtr dst, IntPtr a, IntPtr b, int mode);
        [DllImport("user32.dll")] public static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool redraw);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int index);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
        private static extern int GetWindowLong32(IntPtr hWnd, int index);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int index, IntPtr value);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongW")]
        private static extern int SetWindowLong32(IntPtr hWnd, int index, int value);

        public static int GetExStyle(IntPtr hwnd)
        {
            return IntPtr.Size == 8
                ? (int)(long)GetWindowLongPtr64(hwnd, GWL_EXSTYLE)
                : GetWindowLong32(hwnd, GWL_EXSTYLE);
        }

        public static void SetExStyle(IntPtr hwnd, int style)
        {
            if (IntPtr.Size == 8) SetWindowLongPtr64(hwnd, GWL_EXSTYLE, new IntPtr(style));
            else SetWindowLong32(hwnd, GWL_EXSTYLE, style);
        }

        public static void AddExStyle(IntPtr hwnd, int add)
        {
            SetExStyle(hwnd, GetExStyle(hwnd) | add);
        }

        public static void RemoveExStyle(IntPtr hwnd, int remove)
        {
            SetExStyle(hwnd, GetExStyle(hwnd) & ~remove);
        }

        public static void TryEnableDpiAwareness()
        {
            try { if (SetProcessDpiAwarenessContext(new IntPtr(-4))) return; } catch (Exception) { }
            try { if (SetProcessDPIAware()) return; } catch (Exception) { }
        }

        [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr ctx);

        public static void GetVirtualScreen(out int x, out int y, out int w, out int h)
        {
            IntPtr old = IntPtr.Zero;
            try { old = SetThreadDpiAwarenessContext(new IntPtr(-4)); } catch (Exception) { }
            try
            {
                x = GetSystemMetrics(SM_XVIRTUALSCREEN);
                y = GetSystemMetrics(SM_YVIRTUALSCREEN);
                w = GetSystemMetrics(SM_CXVIRTUALSCREEN);
                h = GetSystemMetrics(SM_CYVIRTUALSCREEN);
            }
            finally
            {
                if (old != IntPtr.Zero) { try { SetThreadDpiAwarenessContext(old); } catch (Exception) { } }
            }
        }

        public static int ScreenDpi()
        {
            IntPtr old = IntPtr.Zero;
            try { old = SetThreadDpiAwarenessContext(new IntPtr(-4)); } catch (Exception) { }
            try
            {
                IntPtr dc = GetDC(IntPtr.Zero);
                if (dc != IntPtr.Zero)
                {
                    try
                    {
                        int value = GetDeviceCaps(dc, LOGPIXELSX);
                        if (value >= 96) return value;
                    }
                    finally { ReleaseDC(IntPtr.Zero, dc); }
                }
            }
            catch (Exception) { }
            finally
            {
                if (old != IntPtr.Zero) { try { SetThreadDpiAwarenessContext(old); } catch (Exception) { } }
            }
            return 96;
        }

        public const int ATTACH_PARENT_PROCESS = -1;

        [DllImport("kernel32.dll")]
        public static extern bool AttachConsole(int processId);

        [DllImport("kernel32.dll")]
        public static extern bool AllocConsole();
    }

    internal class Stroke
    {
        public string Tool;
        public Color Color;
        public int Size;
        public int Seed;
        public string Effect = "";
        public string Text;
        public readonly List<PointF> Points = new List<PointF>();

        public Stroke(string tool, Color color, int size, int seed)
        {
            Tool = tool;
            Color = color;
            Size = size;
            Seed = seed;
        }
    }

    internal class OverlayForm : Form
    {
        public const int HOTKEY_TOGGLE = 1, HOTKEY_UNDO = 2, HOTKEY_CLEAR = 3,
                         HOTKEY_SAVE = 4, HOTKEY_PANEL = 5, HOTKEY_QUIT = 6;

        private int vx, vy, vw, vh;

        private Bitmap committed;
        private Bitmap frame;
        private Bitmap scratch;
        private Graphics frameG;

        private IntPtr hdcScreen, hdcMem, hDib, hDibOld;
        private IntPtr dibBits;

        private readonly List<Stroke> strokes = new List<Stroke>();
        private readonly Stack<Action> undo = new Stack<Action>();
        private Stroke current;
        private bool erasing;
        private Rectangle shapeBounds = Rectangle.Empty;

        private Rectangle ringRect = Rectangle.Empty;
        private PointF mousePos;
        private bool mouseInside;
        private Timer regionTimer;
        private Timer metricsTimer;

        private bool drawMode;
        private readonly Dictionary<string, Rectangle> holes = new Dictionary<string, Rectangle>();
        private readonly bool noActivate;

        private PointF textAt;
        private Bitmap measureBmp;
        private Graphics measureG;
        private readonly Dictionary<int, Font> textFonts = new Dictionary<int, Font>();
        private static readonly StringFormat TextFormat = (StringFormat)StringFormat.GenericTypographic.Clone();

        public string Tool = "pen";
        public Color BrushColor = Color.FromArgb(255, 59, 48);
        public int BrushSize = 6;
        public string ShapeEffect = "none";

        public event EventHandler StrokesChanged;
        public event EventHandler DrawModeChanged;
        public event Action<Bitmap> RegionCaptured;

        public bool DrawMode { get { return drawMode; } }
        public int StrokeCount { get { return strokes.Count; } }
        public int ViewWidth { get { return vw; } }
        public int ViewHeight { get { return vh; } }

        public OverlayForm(int x, int y, int w, int h, bool noActivate)
        {
            vx = x; vy = y; vw = w; vh = h;
            this.noActivate = noActivate;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            Bounds = new Rectangle(x, y, w, h);
            TopMost = true;
            BackColor = Color.Black;
            Text = "ScreenPaint";
            SetStyle(ControlStyles.Opaque | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            DoubleBuffered = false;

            committed = new Bitmap(w, h, PixelFormat.Format32bppPArgb);
            scratch = new Bitmap(w, h, PixelFormat.Format32bppPArgb);
            AutoScaleMode = AutoScaleMode.None;
            CreateSurface();
        }

        private void CreateSurface()
        {
            hdcScreen = Native.GetDC(IntPtr.Zero);
            Native.BITMAPINFO bmi = new Native.BITMAPINFO();
            bmi.bmiHeader.biSize = Marshal.SizeOf(typeof(Native.BITMAPINFOHEADER));
            bmi.bmiHeader.biWidth = vw;
            bmi.bmiHeader.biHeight = -vh;          // top-down
            bmi.bmiHeader.biPlanes = 1;
            bmi.bmiHeader.biBitCount = 32;
            bmi.bmiHeader.biCompression = 0;       // BI_RGB
            IntPtr bits;
            hDib = Native.CreateDIBSection(hdcScreen, ref bmi, 0, out bits, IntPtr.Zero, 0);
            if (hDib == IntPtr.Zero) throw new Exception("CreateDIBSection failed");
            dibBits = bits;
            hdcMem = Native.CreateCompatibleDC(hdcScreen);
            hDibOld = Native.SelectObject(hdcMem, hDib);
            frame = new Bitmap(vw, vh, vw * 4, PixelFormat.Format32bppPArgb, bits);
            frameG = Graphics.FromImage(frame);
        }

        private void ReleaseSurface()
        {
            if (frameG != null) { frameG.Dispose(); frameG = null; }
            if (frame != null) { frame.Dispose(); frame = null; }
            if (hdcMem != IntPtr.Zero && hDibOld != IntPtr.Zero) Native.SelectObject(hdcMem, hDibOld);
            hDibOld = IntPtr.Zero;
            if (hDib != IntPtr.Zero) { Native.DeleteObject(hDib); hDib = IntPtr.Zero; }
            if (hdcMem != IntPtr.Zero) { Native.DeleteDC(hdcMem); hdcMem = IntPtr.Zero; }
            if (hdcScreen != IntPtr.Zero) { Native.ReleaseDC(IntPtr.Zero, hdcScreen); hdcScreen = IntPtr.Zero; }
            dibBits = IntPtr.Zero;
        }

        public bool ApplyScreenBounds(int x, int y, int w, int h)
        {
            if (w <= 0 || h <= 0) return false;
            if (x == vx && y == vy && w == vw && h == vh) return false;

            double sx = (double)w / vw, sy = (double)h / vh;
            for (int i = 0; i < strokes.Count; i++)
            {
                Stroke s = strokes[i];
                for (int k = 0; k < s.Points.Count; k++)
                {
                    PointF p = s.Points[k];
                    s.Points[k] = new PointF((float)(p.X * sx), (float)(p.Y * sy));
                }
            }

            current = null;
            erasing = false;
            mouseInside = false;
            ringRect = Rectangle.Empty;

            ReleaseSurface();
            if (committed != null) { committed.Dispose(); committed = null; }
            if (scratch != null) { scratch.Dispose(); scratch = null; }

            vx = x; vy = y; vw = w; vh = h;

            committed = new Bitmap(vw, vh, PixelFormat.Format32bppPArgb);
            scratch = new Bitmap(vw, vh, PixelFormat.Format32bppPArgb);
            CreateSurface();

            if (IsHandleCreated)
            {
                SetBounds(vx, vy, vw, vh);
                UpdateRegion();
            }
            RebuildCommitted();
            PresentAll();
            Program.Log(string.Format("resize: {0},{1} {2}x{3}", vx, vy, vw, vh));
            return true;
        }

        public bool RefreshScreenMetrics()
        {
            int nx, ny, nw, nh;
            Native.GetVirtualScreen(out nx, out ny, out nw, out nh);
            if (nw <= 0 || nh <= 0) return false;
            return ApplyScreenBounds(nx, ny, nw, nh);
        }

        private void ScheduleMetricsCheck()
        {
            if (metricsTimer == null)
            {
                metricsTimer = new Timer();
                metricsTimer.Interval = 300;
                metricsTimer.Tick += delegate
                {
                    metricsTimer.Stop();
                    try { RefreshScreenMetrics(); }
                    catch (Exception ex) { Program.Log("resize: " + ex); }
                };
            }
            metricsTimer.Stop();
            metricsTimer.Start();
        }

        public bool DebugSetScreenBounds(int x, int y, int w, int h)
        {
            return ApplyScreenBounds(x, y, w, h);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            Native.AddExStyle(Handle, Native.WS_EX_LAYERED | Native.WS_EX_TOOLWINDOW);
            ApplyHitMode();

            uint mods = Native.MOD_CONTROL | Native.MOD_ALT | Native.MOD_NOREPEAT;
            Program.Log(string.Format("hotkeys: D={0} Z={1} X={2} S={3} H={4} Q={5}",
                Native.RegisterHotKey(Handle, HOTKEY_TOGGLE, mods, (uint)Keys.D),
                Native.RegisterHotKey(Handle, HOTKEY_UNDO, mods, (uint)Keys.Z),
                Native.RegisterHotKey(Handle, HOTKEY_CLEAR, mods, (uint)Keys.X),
                Native.RegisterHotKey(Handle, HOTKEY_SAVE, mods, (uint)Keys.S),
                Native.RegisterHotKey(Handle, HOTKEY_PANEL, mods, (uint)Keys.H),
                Native.RegisterHotKey(Handle, HOTKEY_QUIT, mods, (uint)Keys.Q)));
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            for (int id = HOTKEY_TOGGLE; id <= HOTKEY_QUIT; id++) Native.UnregisterHotKey(Handle, id);
            base.OnHandleDestroyed(e);
        }

        protected override void OnPaintBackground(PaintEventArgs e) {}
        protected override void OnPaint(PaintEventArgs e) { }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= Native.WS_EX_LAYERED | Native.WS_EX_TOOLWINDOW;
                if (noActivate) cp.ExStyle |= Native.WS_EX_NOACTIVATE;
                return cp;
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Native.WM_HOTKEY)
            {
                int id = m.WParam.ToInt32();
                OnHotkey(id);
                return;
            }
            if (m.Msg == Native.WM_DPICHANGED || m.Msg == Native.WM_DISPLAYCHANGE)
            {
                base.WndProc(ref m);
                try { RefreshScreenMetrics(); }
                catch (Exception ex) { Program.Log("resize: " + ex); }
                ScheduleMetricsCheck();
                return;
            }
            base.WndProc(ref m);
        }

        public event EventHandler<string> HotkeyPressed;

        private void OnHotkey(int id)
        {
            Program.Log("hotkey id=" + id);
            EventHandler<string> h = HotkeyPressed;
            if (h == null) return;
            switch (id)
            {
                case HOTKEY_TOGGLE: h(this, "toggle"); break;
                case HOTKEY_UNDO: h(this, "undo"); break;
                case HOTKEY_CLEAR: h(this, "clear"); break;
                case HOTKEY_SAVE: h(this, "save"); break;
                case HOTKEY_PANEL: h(this, "panel"); break;
                case HOTKEY_QUIT: h(this, "quit"); break;
            }
        }

        private void ApplyHitMode()
        {
            if (!IsHandleCreated) return;
            if (drawMode) Native.RemoveExStyle(Handle, Native.WS_EX_TRANSPARENT);
            else Native.AddExStyle(Handle, Native.WS_EX_TRANSPARENT);
        }

        public void SetDrawMode(bool on)
        {
            if (drawMode == on) return;
            Program.Log("draw mode = " + on);
            drawMode = on;
            ApplyHitMode();
            if (drawMode)
            {
                Cursor = Cursors.Cross;
                mouseInside = false;
            }
            else
            {
                Cursor = Cursors.Default;
                ringRect = Rectangle.Empty;
                mouseInside = false;
            }
            RebuildCommitted();
            PresentAll();
            EventHandler h = DrawModeChanged;
            if (h != null) h(this, EventArgs.Empty);
        }

        public void SetToolbarHole(Rectangle screenRect)
        {
            SetHole("panel", screenRect);
        }

        public void ClearToolbarHole()
        {
            ClearHole("panel");
        }

        public void SetHole(string name, Rectangle screenRect)
        {
            Rectangle old;
            if (holes.TryGetValue(name, out old) && old == screenRect) return;
            holes[name] = screenRect;
            UpdateRegion();
            ScheduleRebuild();
        }

        public void ClearHole(string name)
        {
            if (!holes.Remove(name)) return;
            UpdateRegion();
            ScheduleRebuild();
        }

        private void ScheduleRebuild()
        {
            if (regionTimer == null)
            {
                regionTimer = new Timer();
                regionTimer.Interval = 120;
                regionTimer.Tick += delegate
                {
                    regionTimer.Stop();
                    RebuildCommitted();
                    PresentAll();
                };
            }
            regionTimer.Stop();
            regionTimer.Start();
        }

        private void UpdateRegion()
        {
            if (!IsHandleCreated) return;
            IntPtr rgn = Native.CreateRectRgn(0, 0, vw, vh);
            foreach (Rectangle h in holes.Values)
            {
                IntPtr cut = Native.CreateRectRgn(
                    h.Left - vx, h.Top - vy,
                    h.Right - vx, h.Bottom - vy);
                Native.CombineRgn(rgn, rgn, cut, Native.RGN_DIFF);
                Native.DeleteObject(cut);
            }
            Native.SetWindowRgn(Handle, rgn, true);
        }

        private Rectangle Full()
        {
            return new Rectangle(0, 0, vw, vh);
        }

        private void ApplyVeil(Graphics g)
        {
            if (!drawMode) return;

            using (SolidBrush b = new SolidBrush(Color.FromArgb(1, 0, 0, 0)))
                g.FillRectangle(b, 0, 0, vw, vh);
        }

        public void RebuildCommitted()
        {
            using (Graphics g = Graphics.FromImage(committed))
            {
                g.CompositingMode = CompositingMode.SourceCopy;
                g.Clear(Color.Transparent);
                g.CompositingMode = CompositingMode.SourceOver;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                ApplyVeil(g);
                for (int i = 0; i < strokes.Count; i++) PaintStroke(g, strokes[i], Full());
            }
        }

        public void PresentAll()
        {
            Graphics g = frameG;
            g.ResetClip();
            g.CompositingMode = CompositingMode.SourceCopy;
            g.DrawImage(committed, 0, 0, vw, vh);
            g.CompositingMode = CompositingMode.SourceOver;
            if (current != null) PaintStroke(g, current, Full());
            if (RingVisible) DrawRing(g, mousePos);
            Push();
        }

        public int SurfaceAlphaAt(int x, int y)
        {
            if (x < 0 || y < 0 || x >= vw || y >= vh) return -1;
            return Marshal.ReadByte(dibBits, (y * vw + x) * 4 + 3);
        }

        private void Push()
        {
            Native.POINT dst = new Native.POINT(vx, vy);
            Native.POINT src = new Native.POINT(0, 0);
            Native.SIZE size = new Native.SIZE(vw, vh);
            Native.BLENDFUNCTION bf = new Native.BLENDFUNCTION();
            bf.BlendOp = Native.AC_SRC_OVER;
            bf.BlendFlags = 0;
            bf.SourceConstantAlpha = 255;
            bf.AlphaFormat = Native.AC_SRC_ALPHA;
            Native.UpdateLayeredWindow(Handle, hdcScreen, ref dst, ref size, hdcMem, ref src, 0, ref bf, Native.ULW_ALPHA);
        }

        internal static float ToolWidth(string tool, int size)
        {
            if (tool == "marker") return size * 1.6f;
            if (tool == "highlighter") return size * 2.4f;
            if (tool == "text") return TextPixelSize(size);
            if (tool == "shot") return 2f;
            return size;
        }

        internal static int TextPixelSize(int size)
        {
            return Math.Max(12, (int)Math.Round(size * 3.5));
        }

        private static int ToolAlpha(string tool)
        {
            if (tool == "marker") return 128;
            if (tool == "highlighter") return 77;
            return 255;
        }

        private static bool IsShapeTool(string tool)
        {
            return tool == "line" || tool == "arrow" || tool == "rect" ||
                   tool == "ellipse" || tool == "triangle";
        }

        internal static bool IsClosedShape(string tool)
        {
            return tool == "rect" || tool == "ellipse" || tool == "triangle";
        }

        private static bool IsDragTool(string tool)
        {
            return IsShapeTool(tool) || tool == "shot";
        }

        public const int LightAlpha = 90;
        public const int DimAlpha = 128;

        private static GraphicsPath ShapePath(Stroke s)
        {
            RectangleF box = ShapeBox(s);
            GraphicsPath path = new GraphicsPath();
            if (s.Tool == "ellipse") path.AddEllipse(box);
            else if (s.Tool == "triangle") path.AddPolygon(TrianglePoints(box));
            else if (s.Tool == "rect") path.AddRectangle(box);
            else
            {
                PointF a = s.Points[0];
                PointF b = s.Points[s.Points.Count - 1];
                path.AddLine(a, b);
            }
            return path;
        }

        private void PaintShapeEffect(Graphics g, Stroke s)
        {
            if (s.Effect == null || s.Effect.Length == 0) return;

            if (s.Effect == "light")
            {
                using (GraphicsPath path = ShapePath(s))
                using (SolidBrush b = new SolidBrush(Color.FromArgb(LightAlpha, 255, 255, 255)))
                    g.FillPath(b, path);
                return;
            }

            using (GraphicsPath path = new GraphicsPath(FillMode.Alternate))
            {
                path.AddRectangle(Full());
                using (GraphicsPath inner = ShapePath(s)) path.AddPath(inner, false);
                using (SolidBrush b = new SolidBrush(Color.FromArgb(DimAlpha, 0, 0, 0)))
                    g.FillPath(b, path);
            }
        }

        private static bool ShapeHitInside(Stroke s, float x, float y)
        {
            if (s.Effect == null || s.Effect.Length == 0) return false;
            if (!IsClosedShape(s.Tool) || s.Points.Count < 2) return false;
            using (GraphicsPath path = ShapePath(s))
                return path.IsVisible(x, y);
        }

        private Font TextFont(int size)
        {
            int px = TextPixelSize(size);
            Font f;
            if (textFonts.TryGetValue(px, out f)) return f;
            f = new Font("Segoe UI", px, FontStyle.Regular, GraphicsUnit.Pixel);
            textFonts[px] = f;
            return f;
        }

        private Graphics MeasureGraphics()
        {
            if (measureG == null)
            {
                measureBmp = new Bitmap(1, 1);
                measureG = Graphics.FromImage(measureBmp);
            }
            return measureG;
        }

        private RectangleF TextBounds(Stroke s)
        {
            PointF at = s.Points.Count > 0 ? s.Points[0] : PointF.Empty;
            if (string.IsNullOrEmpty(s.Text)) return new RectangleF(at.X, at.Y, 2f, 2f);
            SizeF size = MeasureGraphics().MeasureString(s.Text, TextFont(s.Size), PointF.Empty, TextFormat);
            return new RectangleF(at.X, at.Y, size.Width, size.Height);
        }

        private float TextLineHeight(int size)
        {
            return MeasureGraphics().MeasureString("Ag", TextFont(size), PointF.Empty, TextFormat).Height;
        }

        private void PaintTextStroke(Graphics g, Stroke s)
        {
            if (s.Points.Count == 0 || string.IsNullOrEmpty(s.Text)) return;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            // ClearType на прозрачном слое даёт цветную бахрому — нужен обычный AA.
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            using (SolidBrush b = new SolidBrush(s.Color))
                g.DrawString(s.Text, TextFont(s.Size), b, s.Points[0], TextFormat);
        }

        private static readonly Font ShotFont = new Font("Segoe UI", 14f, FontStyle.Regular, GraphicsUnit.Pixel);

        private void PaintShotBand(Graphics g, Stroke s)
        {
            RectangleF box = ShapeBox(s);
            if (box.Width < 1f && box.Height < 1f) return;

            g.SmoothingMode = SmoothingMode.None;
            using (Pen back = new Pen(Color.FromArgb(190, 0, 0, 0), 3f))
                g.DrawRectangle(back, box.X, box.Y, box.Width, box.Height);
            using (Pen dash = new Pen(Color.White, 1f))
            {
                dash.DashStyle = DashStyle.Custom;
                dash.DashPattern = new float[] { 4f, 4f };
                g.DrawRectangle(dash, box.X, box.Y, box.Width, box.Height);
            }

            string label = string.Format("{0}×{1}", (int)Math.Round(box.Width), (int)Math.Round(box.Height));
            Size text = TextRenderer.MeasureText(label, ShotFont);
            if (box.Width < text.Width + 12f || box.Height < text.Height + 12f) return;

            RectangleF tag = new RectangleF(box.Left + 4f, box.Top + 4f, text.Width + 8f, text.Height + 2f);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (SolidBrush bg = new SolidBrush(Color.FromArgb(200, 0, 0, 0)))
                g.FillRectangle(bg, tag);
            using (SolidBrush fg = new SolidBrush(Color.White))
                g.DrawString(label, ShotFont, fg, tag.Left + 4f, tag.Top + 1f, TextFormat);
        }

        private void PaintStroke(Graphics g, Stroke s, Rectangle clip)
        {
            if (s.Points.Count == 0) return;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.None;

            if (s.Tool == "spray")
            {
                PaintSpray(g, s, clip);
                return;
            }

            if (s.Tool == "text")
            {
                PaintTextStroke(g, s);
                return;
            }

            if (s.Tool == "shot")
            {
                PaintShotBand(g, s);
                return;
            }

            float width = ToolWidth(s.Tool, s.Size);
            int alpha = ToolAlpha(s.Tool);
            Color color = s.Color;
            bool shape = IsShapeTool(s.Tool);

            bool shortStroke = !shape && (s.Points.Count == 1 || (s.Points.Count == 2 &&
                Math.Abs(s.Points[0].X - s.Points[1].X) < 0.01f && Math.Abs(s.Points[0].Y - s.Points[1].Y) < 0.01f));

            Action<Graphics> painter = delegate(Graphics t)
            {
                if (shape)
                {
                    PaintShapeEffect(t, s);
                    using (Pen p = MakePen(color, width)) DrawShape(t, s, p);
                }
                else if (shortStroke)
                {
                    using (SolidBrush b = new SolidBrush(color))
                        t.FillEllipse(b, s.Points[0].X - width / 2f, s.Points[0].Y - width / 2f, width, width);
                }
                else
                {
                    using (Pen p = MakePen(color, width))
                        t.DrawLines(p, s.Points.ToArray());
                }
            };

            if (alpha == 255)
            {
                painter(g);
                return;
            }

            Rectangle dirty = StrokeBounds(s, width);
            dirty.Intersect(clip);
            if (dirty.Width <= 0 || dirty.Height <= 0) return;

            using (Graphics sg = Graphics.FromImage(scratch))
            {
                sg.ResetClip();
                sg.SetClip(dirty);
                sg.CompositingMode = CompositingMode.SourceCopy;
                using (SolidBrush clear = new SolidBrush(Color.Transparent))
                    sg.FillRectangle(clear, dirty);
                sg.CompositingMode = CompositingMode.SourceOver;
                sg.SmoothingMode = SmoothingMode.AntiAlias;
                painter(sg);
            }
            float a = alpha / 255f;
            ColorMatrix cm = new ColorMatrix(new float[][] {
                new float[] { 1, 0, 0, 0, 0 },
                new float[] { 0, 1, 0, 0, 0 },
                new float[] { 0, 0, 1, 0, 0 },
                new float[] { 0, 0, 0, a, 0 },
                new float[] { 0, 0, 0, 0, 1 }
            });
            using (ImageAttributes ia = new ImageAttributes())
            {
                ia.SetColorMatrix(cm);
                InterpolationMode old = g.InterpolationMode;
                g.InterpolationMode = InterpolationMode.NearestNeighbor;
                g.DrawImage(scratch, dirty, dirty.Left, dirty.Top, dirty.Width, dirty.Height,
                            GraphicsUnit.Pixel, ia);
                g.InterpolationMode = old;
            }
        }

        private static RectangleF ShapeBox(Stroke s)
        {
            PointF a = s.Points[0];
            PointF b = s.Points[s.Points.Count - 1];
            return new RectangleF(
                Math.Min(a.X, b.X), Math.Min(a.Y, b.Y),
                Math.Abs(a.X - b.X), Math.Abs(a.Y - b.Y));
        }

        private static PointF[] TrianglePoints(RectangleF box)
        {
            return new PointF[] {
                new PointF(box.Left + box.Width / 2f, box.Top),
                new PointF(box.Right, box.Bottom),
                new PointF(box.Left, box.Bottom)
            };
        }

        private static void DrawShape(Graphics g, Stroke s, Pen pen)
        {
            PointF a = s.Points[0];
            PointF b = s.Points[s.Points.Count - 1];
            switch (s.Tool)
            {
                case "line":
                    g.DrawLine(pen, a, b);
                    break;
                case "arrow":
                    DrawArrow(g, a, b, pen);
                    break;
                case "rect":
                {
                    RectangleF box = ShapeBox(s);
                    g.DrawRectangle(pen, box.X, box.Y, box.Width, box.Height);
                    break;
                }
                case "ellipse":
                    g.DrawEllipse(pen, ShapeBox(s));
                    break;
                default:
                    g.DrawPolygon(pen, TrianglePoints(ShapeBox(s)));
                    break;
            }
        }

        private static void DrawArrow(Graphics g, PointF a, PointF b, Pen pen)
        {
            float dx = b.X - a.X, dy = b.Y - a.Y;
            float len = (float)Math.Sqrt(dx * dx + dy * dy);
            if (len < 0.5f) return;

            g.DrawLine(pen, a, b);

            float head = Math.Min(Math.Max(10f, pen.Width * 3f), len);
            double ang = Math.Atan2(dy, dx);
            const double spread = 0.42;                     // ~24 градуса
            PointF p1 = new PointF(
                b.X - (float)(head * Math.Cos(ang - spread)),
                b.Y - (float)(head * Math.Sin(ang - spread)));
            PointF p2 = new PointF(
                b.X - (float)(head * Math.Cos(ang + spread)),
                b.Y - (float)(head * Math.Sin(ang + spread)));
            PointF[] headPts = new PointF[] { b, p1, p2 };

            using (SolidBrush hb = new SolidBrush(pen.Color))
                g.FillPolygon(hb, headPts);
            using (Pen hp = MakePen(pen.Color, pen.Width))
                g.DrawPolygon(hp, headPts);
        }

        private static bool ShiftDown()
        {
            try { return (Native.GetAsyncKeyState(Native.VK_SHIFT) & 0x8000) != 0; }
            catch (Exception) { return false; }
        }

        private static PointF ConstrainShape(PointF a, PointF b, string tool)
        {
            if (!ShiftDown()) return b;

            if (tool == "line" || tool == "arrow")
            {
                double ang = Math.Atan2(b.Y - a.Y, b.X - a.X);
                double step = Math.PI / 4.0;
                double snap = Math.Round(ang / step) * step;
                double len = Math.Sqrt((b.X - a.X) * (b.X - a.X) + (b.Y - a.Y) * (b.Y - a.Y));
                return new PointF(a.X + (float)(len * Math.Cos(snap)), a.Y + (float)(len * Math.Sin(snap)));
            }

            float dx = b.X - a.X, dy = b.Y - a.Y;
            float size = Math.Min(Math.Abs(dx), Math.Abs(dy));
            return new PointF(a.X + Math.Sign(dx) * size, a.Y + Math.Sign(dy) * size);
        }

        private static Pen MakePen(Color color, float width)
        {
            Pen p = new Pen(color, width);
            p.StartCap = LineCap.Round;
            p.EndCap = LineCap.Round;
            p.LineJoin = LineJoin.Round;
            p.MiterLimit = 2f;
            return p;
        }

        private void PaintSpray(Graphics g, Stroke s, Rectangle clip)
        {
            float radius = Math.Max(3f, s.Size);
            int count = Math.Max(4, s.Size + 2);
            float dot = 0.9f + s.Size / 10f;
            Random rnd = new Random(s.Seed);
            using (SolidBrush b = new SolidBrush(Color.FromArgb(217, s.Color)))
            {
                for (int i = 0; i < s.Points.Count; i++)
                {
                    PointF pt = s.Points[i];
                    for (int k = 0; k < count; k++)
                    {
                        double ang = rnd.NextDouble() * Math.PI * 2.0;
                        double r = radius * Math.Sqrt(rnd.NextDouble());
                        float px = pt.X + (float)(r * Math.Cos(ang));
                        float py = pt.Y + (float)(r * Math.Sin(ang));
                        if (px + dot < clip.Left - 2 || px - dot > clip.Right + 2 ||
                            py + dot < clip.Top - 2 || py - dot > clip.Bottom + 2) continue;
                        g.FillEllipse(b, px - dot, py - dot, dot * 2f, dot * 2f);
                    }
                }
            }
        }

        private Rectangle StrokeBounds(Stroke s, float width)
        {
            if (s.Points.Count == 0) return Rectangle.Empty;

            if (s.Tool == "text")
            {
                RectangleF tb = TextBounds(s);
                return Rectangle.FromLTRB(
                    (int)Math.Floor(tb.Left) - 3, (int)Math.Floor(tb.Top) - 3,
                    (int)Math.Ceiling(tb.Right) + 4, (int)Math.Ceiling(tb.Bottom) + 4);
            }

            float minX = float.MaxValue, minY = float.MaxValue, maxX = float.MinValue, maxY = float.MinValue;
            for (int i = 0; i < s.Points.Count; i++)
            {
                PointF p = s.Points[i];
                if (p.X < minX) minX = p.X;
                if (p.Y < minY) minY = p.Y;
                if (p.X > maxX) maxX = p.X;
                if (p.Y > maxY) maxY = p.Y;
            }
            float pad = width / 2f + 3f;
            if (s.Tool == "spray") pad = Math.Max(3f, s.Size) + 4f;
            else if (s.Tool == "arrow")
                pad = Math.Max(pad, Math.Max(10f, width * 3f) * 0.5f + width / 2f + 2f);
            return Rectangle.FromLTRB(
                (int)Math.Floor(minX - pad), (int)Math.Floor(minY - pad),
                (int)Math.Ceiling(maxX + pad) + 1, (int)Math.Ceiling(maxY + pad) + 1);
        }

        private void DrawRing(Graphics g, PointF p)
        {
            float r = (Tool == "eraser" ? BrushSize : ToolWidth(Tool, BrushSize)) / 2f;
            if (r < 2f) r = 2f;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (Pen pen = new Pen(Color.FromArgb(210, 0, 229, 255), 1f))
                g.DrawEllipse(pen, p.X - r, p.Y - r, r * 2f, r * 2f);
        }

        private bool RingVisible
        {
            get { return drawMode && mouseInside && Tool != "shot"; }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (!drawMode) { base.OnMouseDown(e); return; }
            if (e.Button == MouseButtons.Right) { Undo(); return; }
            if (e.Button != MouseButtons.Left) return;
            mouseInside = true;
            mousePos = new PointF(e.X, e.Y);
            if (Tool == "eraser")
            {
                erasing = true;
                EraseAt(e.X, e.Y);
                return;
            }
            if (Tool == "text")
            {
                textAt = mousePos;
                return;
            }
            Stroke s = new Stroke(Tool, BrushColor, BrushSize, Environment.TickCount & 0x7FFFFFFF ^ (strokes.Count * 7919));
            s.Points.Add(mousePos);
            bool shape = IsDragTool(Tool);
            if (shape) s.Points.Add(mousePos);
            if (IsClosedShape(Tool) && ShapeEffect != "none") s.Effect = ShapeEffect;
            strokes.Add(s);
            current = s;
            Stroke added = s;
            if (!shape) undo.Push(delegate { strokes.Remove(added); });
            Rectangle b = StrokeBounds(s, ToolWidth(Tool, BrushSize));
            if (shape) shapeBounds = b;
            RenderLive(b);
            RaiseStrokesChanged();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (!drawMode) { base.OnMouseMove(e); return; }
            mouseInside = true;
            PointF p = new PointF(e.X, e.Y);

            if (erasing)
            {
                mousePos = p;
                EraseAt(e.X, e.Y);
                RenderLive(Rectangle.Empty);
                return;
            }

            if (current == null)
            {
                mousePos = p;
                RenderLive(Rectangle.Empty);
                return;
            }

            if (current.Tool == "text") return;

            if (IsDragTool(current.Tool))
            {
                PointF end = IsShapeTool(current.Tool) ? ConstrainShape(current.Points[0], p, current.Tool) : p;
                mousePos = end;
                Rectangle before = shapeBounds;
                current.Points[1] = end;
                Rectangle after = StrokeBounds(current, ToolWidth(current.Tool, current.Size));
                shapeBounds = after;
                RenderLive(Union(before, after));
                return;
            }

            PointF last = current.Points[current.Points.Count - 1];
            mousePos = p;
            if (Math.Abs(last.X - p.X) + Math.Abs(last.Y - p.Y) < 2f)
            {
                RenderLive(Rectangle.Empty);
                return;
            }

            Rectangle seg = Rectangle.FromLTRB(
                (int)Math.Floor(Math.Min(last.X, p.X)), (int)Math.Floor(Math.Min(last.Y, p.Y)),
                (int)Math.Ceiling(Math.Max(last.X, p.X)) + 1, (int)Math.Ceiling(Math.Max(last.Y, p.Y)) + 1);
            float pad = Math.Max(ToolWidth(current.Tool, current.Size), Math.Max(3f, current.Size)) + 6f;
            seg.Inflate((int)pad, (int)pad);
            current.Points.Add(p);
            RenderLive(seg);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (!drawMode) { base.OnMouseUp(e); return; }
            if (e.Button != MouseButtons.Left) return;
            mouseInside = true;
            mousePos = new PointF(e.X, e.Y);
            if (Tool == "text")
            {
                BeginTextInput(textAt);
                return;
            }
            if (current != null && current.Tool == "shot")
            {
                Stroke band = current;
                current = null;
                shapeBounds = Rectangle.Empty;
                strokes.Remove(band);
                RebuildCommitted();
                PresentAll();
                RaiseStrokesChanged();
                RectangleF box = ShapeBox(band);
                if (box.Width >= 4f && box.Height >= 4f) HandleShot(box);
                else Program.Log("shot: too small");
                return;
            }
            bool had = current != null;
            if (current != null && IsShapeTool(current.Tool))
            {
                Stroke sh = current;
                float w = Math.Abs(sh.Points[0].X - sh.Points[1].X);
                float hh = Math.Abs(sh.Points[0].Y - sh.Points[1].Y);
                if (w < 2f && hh < 2f)
                {
                    strokes.Remove(sh);
                    had = false;
                }
                else
                {
                    undo.Push(delegate { strokes.Remove(sh); });
                }
            }
            current = null;
            shapeBounds = Rectangle.Empty;
            erasing = false;
            mousePos = new PointF(e.X, e.Y);
            RebuildCommitted();
            PresentAll();
            if (had) RaiseStrokesChanged();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            if (drawMode && mouseInside)
            {
                mouseInside = false;
                RenderLive(Rectangle.Empty);
            }
            base.OnMouseLeave(e);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            if (drawMode)
            {
                int next = BrushSize + (e.Delta > 0 ? 1 : -1);
                BrushSize = Math.Max(1, Math.Min(60, next));
                mousePos = new PointF(e.X, e.Y);
                RenderLive(Rectangle.Empty);
                RaiseStrokesChanged();
            }
            base.OnMouseWheel(e);
        }

        private void BeginTextInput(PointF at)
        {
            Stroke s = new Stroke("text", BrushColor, BrushSize, Environment.TickCount & 0x7FFFFFFF);
            s.Points.Add(at);
            s.Text = "";
            Point screen = new Point(vx + (int)Math.Round(at.X), vy + (int)Math.Round(at.Y));
            Program.Log(string.Format("text: input at {0},{1}", screen.X, screen.Y));
            BeginInvoke((MethodInvoker)delegate { EditTextStroke(s, screen); });
        }

        private void EditTextStroke(Stroke s, Point screen)
        {
            bool ok;
            using (TextDialog dlg = new TextDialog(Native.ScreenDpi()))
            {
                Program.Log(string.Format("text dialog: {0}x{1}", dlg.Width, dlg.Height));
                dlg.Location = ClampToDesktop(
                    new Point(screen.X, screen.Y + (int)Math.Round(TextLineHeight(s.Size)) + 8), dlg.Size);
                dlg.Typed += delegate(string text)
                {
                    Rectangle before = StrokeBounds(s, 0f);
                    s.Text = text;
                    RenderLive(Union(before, StrokeBounds(s, 0f)));
                };
                current = s;
                ok = dlg.ShowDialog(this) == DialogResult.OK && dlg.Value.Trim().Length > 0;
                if (ok) s.Text = dlg.Value;
            }

            current = null;
            if (ok)
            {
                strokes.Add(s);
                Stroke added = s;
                undo.Push(delegate { strokes.Remove(added); });
                Program.Log("text: added \"" + s.Text + "\"");
                RaiseStrokesChanged();
            }
            else
            {
                Program.Log("text: cancelled");
            }
            RebuildCommitted();
            PresentAll();
        }

        private static Point ClampToDesktop(Point at, Size size)
        {
            int vx0, vy0, vw0, vh0;
            Native.GetVirtualScreen(out vx0, out vy0, out vw0, out vh0);
            if (vw0 <= 0 || vh0 <= 0) return at;
            int x = Math.Max(vx0, Math.Min(at.X, vx0 + vw0 - size.Width));
            int y = Math.Max(vy0, Math.Min(at.Y, vy0 + vh0 - size.Height));
            return new Point(x, y);
        }

        private void RenderLive(Rectangle extraDirty)
        {
            Rectangle dirty = extraDirty;
            if (!ringRect.IsEmpty) dirty = Union(dirty, ringRect);
            Rectangle newRing = RingVisible ? RingBounds(mousePos) : Rectangle.Empty;
            if (!newRing.IsEmpty) dirty = Union(dirty, newRing);
            if (current != null && current.Effect == "dim") dirty = Full();
            if (dirty.IsEmpty) dirty = Full();
            dirty.Inflate(2, 2);
            dirty.Intersect(Full());
            if (dirty.Width <= 0 || dirty.Height <= 0) return;

            Graphics g = frameG;
            g.ResetClip();
            g.SetClip(dirty);
            g.CompositingMode = CompositingMode.SourceCopy;
            g.DrawImage(committed, dirty, dirty, GraphicsUnit.Pixel);
            g.CompositingMode = CompositingMode.SourceOver;
            if (current != null) PaintStroke(g, current, dirty);
            if (!newRing.IsEmpty) DrawRing(g, mousePos);
            ringRect = newRing;
            Push();
        }

        private Rectangle RingBounds(PointF p)
        {
            float r = (Tool == "eraser" ? BrushSize : ToolWidth(Tool, BrushSize)) / 2f;
            if (r < 2f) r = 2f;
            return Rectangle.FromLTRB(
                (int)Math.Floor(p.X - r) - 2, (int)Math.Floor(p.Y - r) - 2,
                (int)Math.Ceiling(p.X + r) + 3, (int)Math.Ceiling(p.Y + r) + 3);
        }

        private static Rectangle Union(Rectangle a, Rectangle b)
        {
            if (a.IsEmpty) return b;
            if (b.IsEmpty) return a;
            return Rectangle.Union(a, b);
        }

        private void EraseAt(float x, float y)
        {
            List<Stroke> hit = new List<Stroke>();
            for (int i = 0; i < strokes.Count; i++)
            {
                Stroke s = strokes[i];
                float reach = BrushSize / 2f + Math.Max(2f, ToolWidth(s.Tool, s.Size) / 2f);
                Rectangle b = StrokeBounds(s, ToolWidth(s.Tool, s.Size));
                if (x < b.Left - reach || x > b.Right + reach || y < b.Top - reach || y > b.Bottom + reach) continue;
                if (DistToStroke(x, y, s) <= reach || ShapeHitInside(s, x, y)) hit.Add(s);
            }
            if (hit.Count == 0) return;
            for (int i = 0; i < hit.Count; i++)
            {
                Stroke s = hit[i];
                int index = strokes.IndexOf(s);
                strokes.Remove(s);
                undo.Push(delegate { strokes.Insert(Math.Min(index, strokes.Count), s); });
            }
            RebuildCommitted();
            PresentAll();
            RaiseStrokesChanged();
        }

        private float DistToStroke(float px, float py, Stroke s)
        {
            List<PointF> pts = s.Points;
            if (pts.Count == 0) return float.MaxValue;

            if (s.Tool == "text")
            {
                RectangleF tb = TextBounds(s);
                float dx = Math.Max(Math.Max(tb.Left - px, px - tb.Right), 0f);
                float dy = Math.Max(Math.Max(tb.Top - py, py - tb.Bottom), 0f);
                return (float)Math.Sqrt(dx * dx + dy * dy);
            }

            if (pts.Count == 1) return Dist(px, py, pts[0].X, pts[0].Y);

            if (IsShapeTool(s.Tool) && pts.Count >= 2)
            {
                RectangleF box = ShapeBox(s);
                switch (s.Tool)
                {
                    case "rect":
                        return Math.Min(
                            Math.Min(SegDist(px, py, box.Left, box.Top, box.Right, box.Top),
                                     SegDist(px, py, box.Right, box.Top, box.Right, box.Bottom)),
                            Math.Min(SegDist(px, py, box.Right, box.Bottom, box.Left, box.Bottom),
                                     SegDist(px, py, box.Left, box.Bottom, box.Left, box.Top)));
                    case "ellipse":
                    {
                        float rx = box.Width / 2f, ry = box.Height / 2f;
                        if (rx < 1f || ry < 1f)
                            return SegDist(px, py, box.Left, box.Top, box.Right, box.Bottom);
                        float u = (px - box.Left - rx) / rx;
                        float v = (py - box.Top - ry) / ry;
                        float k = (float)Math.Sqrt(u * u + v * v);
                        return Math.Abs(k - 1f) * Math.Min(rx, ry);
                    }
                    case "triangle":
                    {
                        PointF[] t = TrianglePoints(box);
                        return Math.Min(
                            SegDist(px, py, t[0].X, t[0].Y, t[1].X, t[1].Y),
                            Math.Min(SegDist(px, py, t[1].X, t[1].Y, t[2].X, t[2].Y),
                                     SegDist(px, py, t[2].X, t[2].Y, t[0].X, t[0].Y)));
                    }
                }
            }

            float best = float.MaxValue;
            for (int i = 1; i < pts.Count; i++)
            {
                float d = SegDist(px, py, pts[i - 1].X, pts[i - 1].Y, pts[i].X, pts[i].Y);
                if (d < best) best = d;
            }
            return best;
        }

        private static float SegDist(float px, float py, float x0, float y0, float x1, float y1)
        {
            float dx = x1 - x0, dy = y1 - y0;
            float len2 = dx * dx + dy * dy;
            float t = len2 <= 0 ? 0f : ((px - x0) * dx + (py - y0) * dy) / len2;
            if (t < 0f) t = 0f;
            if (t > 1f) t = 1f;
            return Dist(px, py, x0 + t * dx, y0 + t * dy);
        }

        private static float Dist(float ax, float ay, float bx, float by)
        {
            float dx = ax - bx, dy = ay - by;
            return (float)Math.Sqrt(dx * dx + dy * dy);
        }

        public void AddStroke(Stroke s)
        {
            strokes.Add(s);
            Stroke added = s;
            undo.Push(delegate { strokes.Remove(added); });
            RaiseStrokesChanged();
        }

        public void Undo()
        {
            if (undo.Count == 0) return;
            undo.Pop()();
            RebuildCommitted();
            PresentAll();
            RaiseStrokesChanged();
        }

        public void ClearAll()
        {
            if (strokes.Count == 0) return;
            List<Stroke> all = new List<Stroke>(strokes);
            strokes.Clear();
            undo.Push(delegate { strokes.AddRange(all); });
            RebuildCommitted();
            PresentAll();
            RaiseStrokesChanged();
        }

        private void RaiseStrokesChanged()
        {
            EventHandler h = StrokesChanged;
            if (h != null) h(this, EventArgs.Empty);
        }

        public void SavePng(string path)
        {
            using (Bitmap bmp = new Bitmap(vw, vh, PixelFormat.Format32bppArgb))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CompositingMode = CompositingMode.SourceOver;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    Rectangle full = Full();
                    for (int i = 0; i < strokes.Count; i++) PaintStroke(g, strokes[i], full);
                }
                bmp.Save(path, ImageFormat.Png);
            }
        }

        public Bitmap CaptureRegion(RectangleF area)
        {
            Rectangle r = Rectangle.FromLTRB(
                (int)Math.Floor(area.Left), (int)Math.Floor(area.Top),
                (int)Math.Ceiling(area.Right), (int)Math.Ceiling(area.Bottom));
            r.Intersect(Full());
            if (r.Width <= 0 || r.Height <= 0) return null;

            try { Native.DwmFlush(); }
            catch (Exception) { }
            System.Threading.Thread.Sleep(50);

            Bitmap bmp = new Bitmap(r.Width, r.Height, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
                g.CopyFromScreen(vx + r.Left, vy + r.Top, 0, 0, new Size(r.Width, r.Height));
            Program.Log(string.Format("shot: {0},{1} {2}x{3}", r.Left, r.Top, r.Width, r.Height));
            return bmp;
        }

        private void HandleShot(RectangleF area)
        {
            Bitmap bmp = CaptureRegion(area);
            if (bmp == null) return;
            try { Clipboard.SetImage(bmp); }
            catch (Exception ex) { Program.Log("shot: clipboard " + ex.Message); }
            try
            {
                Action<Bitmap> h = RegionCaptured;
                if (h != null) h(bmp);
            }
            finally
            {
                bmp.Dispose();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (frameG != null) frameG.Dispose();
                if (frame != null) frame.Dispose();
                if (committed != null) committed.Dispose();
                if (scratch != null) scratch.Dispose();
                if (measureG != null) { measureG.Dispose(); measureG = null; }
                if (measureBmp != null) { measureBmp.Dispose(); measureBmp = null; }
                if (hdcMem != IntPtr.Zero) { Native.SelectObject(hdcMem, hDibOld); Native.DeleteDC(hdcMem); }
                if (hDib != IntPtr.Zero) Native.DeleteObject(hDib);
                if (hdcScreen != IntPtr.Zero) Native.ReleaseDC(IntPtr.Zero, hdcScreen);
            }
            base.Dispose(disposing);
        }
    }

    internal class TextDialog : Form
    {
        private readonly TextBox box;
        private readonly float scale;

        public event Action<string> Typed;

        public string Value { get { return box.Text; } }

        private int S(int value) { return (int)Math.Round(value * scale); }

        private static readonly Color Bg = Color.FromArgb(0x1B, 0x1B, 0x21);
        private static readonly Color Bg2 = Color.FromArgb(0x26, 0x26, 0x2F);
        private static readonly Color Fg = Color.FromArgb(0xE9, 0xE9, 0xF2);
        private static readonly Color Muted = Color.FromArgb(0x9A, 0x9A, 0xAD);
        private static readonly Color Accent = Color.FromArgb(0x3B, 0x82, 0xF6);

        public TextDialog(int dpi)
        {
            scale = Math.Max(96, dpi) / 96f;

            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            AutoScaleMode = AutoScaleMode.None;
            KeyPreview = true;
            BackColor = Bg;
            Text = "signature";

            box = new TextBox();
            box.Location = new Point(S(8), S(8));
            box.Width = S(304);
            box.Font = new Font("Segoe UI", 10f * scale);
            box.Height = box.PreferredHeight;
            box.BorderStyle = BorderStyle.FixedSingle;
            box.BackColor = Bg2;
            box.ForeColor = Fg;
            box.TextChanged += delegate
            {
                Action<string> h = Typed;
                if (h != null) h(box.Text);
            };
            Controls.Add(box);

            int rowY = box.Bottom + S(8);
            ClientSize = new Size(S(320), rowY + S(24) + S(8));

            Label hint = new Label();
            hint.Text = "enter = ok · esc = cancel";
            hint.ForeColor = Muted;
            hint.Font = new Font("Segoe UI", 8f * scale);
            hint.AutoSize = true;
            hint.Location = new Point(S(8), rowY + S(9));
            Controls.Add(hint);

            Button cancel = MakeButton("cancel", S(88));
            cancel.Location = new Point(S(112), rowY);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            Button ok = MakeButton("ok", S(100));
            ok.Location = new Point(S(204), rowY);
            ok.Click += delegate { DialogResult = DialogResult.OK; };
            Controls.Add(ok);

            AcceptButton = ok;
            CancelButton = cancel;
            Shown += delegate
            {
                Activate();
                box.Focus();
            };
        }

        private Button MakeButton(string text, int width)
        {
            Button b = new Button();
            b.Text = text;
            b.Size = new Size(width, S(24));
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = Accent;
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(0x1D, 0x4E, 0xD8);
            b.BackColor = Color.FromArgb(0x33, 0x33, 0x40);
            b.ForeColor = Fg;
            b.Font = new Font("Segoe UI", 9f * scale);
            b.Cursor = Cursors.Hand;
            b.TabStop = false;
            return b;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using (Pen p = new Pen(Accent, 1f))
                e.Graphics.DrawRectangle(p, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }
    }

    internal class Slider : Control
    {
        public int Minimum = 1, Maximum = 60;
        public float UiScale = 1f;
        private int value = 6;
        private bool dragging;

        private int TrackPad { get { return (int)Math.Round(4 * UiScale); } }
        private int KnobR { get { return Math.Max(3, (int)Math.Round(7 * UiScale)); } }
        private int BarH { get { return Math.Max(2, (int)Math.Round(6 * UiScale)); } }

        public event EventHandler ValueChanged;

        public int Value
        {
            get { return value; }
            set
            {
                int v = Math.Max(Minimum, Math.Min(Maximum, value));
                if (v == this.value) return;
                this.value = v;
                Invalidate();
                EventHandler h = ValueChanged;
                if (h != null) h(this, EventArgs.Empty);
            }
        }

        public Slider()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Height = 22;
            Width = 150;
            Cursor = Cursors.Hand;
        }

        private void SetFromX(int x)
        {
            int pad = KnobR;
            float t = Width <= pad * 2 ? 0f : (x - pad) / (float)(Width - pad * 2);
            if (t < 0f) t = 0f;
            if (t > 1f) t = 1f;
            Value = Minimum + (int)Math.Round(t * (Maximum - Minimum));
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            dragging = true;
            SetFromX(e.X);
            base.OnMouseDown(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (dragging) SetFromX(e.X);
            base.OnMouseMove(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            dragging = false;
            base.OnMouseUp(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int cy = Height / 2;
            int pad = TrackPad, bar = BarH, knob = KnobR;
            float t = (value - Minimum) / (float)(Maximum - Minimum);
            using (SolidBrush track = new SolidBrush(Color.FromArgb(0x33, 0x33, 0x40)))
                g.FillRectangle(track, pad, cy - bar / 2, Width - pad * 2, bar);
            using (SolidBrush fill = new SolidBrush(Color.FromArgb(0x3B, 0x82, 0xF6)))
                g.FillRectangle(fill, pad, cy - bar / 2, (int)((Width - pad * 2) * t), bar);
            int kx = pad + (int)((Width - pad * 2) * t) - knob;
            using (SolidBrush knobBrush = new SolidBrush(Color.White))
                g.FillEllipse(knobBrush, kx, cy - knob, knob * 2, knob * 2);
        }
    }

    internal class Palette : Form
    {
        private readonly OverlayForm overlay;

        private readonly List<Panel> swatches = new List<Panel>();
        private readonly Dictionary<string, Button> toolMap = new Dictionary<string, Button>();

        private Slider sizeSlider;
        private Panel colorPreview;
        private Panel brushPreview;
        private Button modeButton;
        private Button effectButton;
        private Label statusLabel;
        private Point dragOffset;
        private bool dragging;
        private int dpi = 96;
        private readonly ToolTip tips = new ToolTip();

        private MiniIcon icon;
        private Timer watch;
        private bool collapsed;
        private bool autoHide;
        private int awayTicks;
        private int modal;
        private DateTime expandedAt = DateTime.MinValue;

        private readonly Font fontBtn = new Font("Segoe UI", 9f);
        private readonly Font fontTitle = new Font("Segoe UI", 10f, FontStyle.Bold);
        private readonly Font fontMode = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        private readonly Font fontStatus = new Font("Consolas", 7.5f);
        private readonly Dictionary<string, Font> fonts = new Dictionary<string, Font>();

        private static readonly string[] ToolKeys = {
            "pen", "marker", "highlighter", "spray", "eraser",
            "line", "arrow", "rect", "ellipse", "triangle",
            "text", "shot"
        };
        private static readonly string[] ToolNames = {
            "pen", "marker", "marker+", "spray", "eraser",
            "line", "arrow", "rect", "ellipse", "triangle",
            "text", "shot"
        };
        private static readonly string[] EffectKeys = { "none", "light", "dim" };
        private static readonly string[] EffectNames = { "effect: none", "effect: light", "effect: dim" };
        private static readonly string[] EffectTips = {
            "no effect — plain outline",
            "highlight: background inside the selection gets lighter",
            "dim: keep the selection as it is, darken the rest of the screen"
        };
        private static readonly Color[] Colors = {
            Color.FromArgb(255, 59, 48), Color.FromArgb(255, 149, 0), Color.FromArgb(255, 214, 10),
            Color.FromArgb(52, 199, 89), Color.FromArgb(0, 199, 190), Color.FromArgb(10, 132, 255),
            Color.FromArgb(94, 92, 230), Color.FromArgb(191, 90, 242), Color.FromArgb(255, 45, 85),
            Color.White, Color.FromArgb(199, 199, 204), Color.FromArgb(142, 142, 147),
            Color.FromArgb(72, 72, 74), Color.FromArgb(28, 28, 30)
        };

        internal static readonly Color Bg = Color.FromArgb(0x1B, 0x1B, 0x21);
        internal static readonly Color Bg2 = Color.FromArgb(0x26, 0x26, 0x2F);
        internal static readonly Color BgHover = Color.FromArgb(0x3A, 0x3A, 0x48);
        internal static readonly Color Fg = Color.FromArgb(0xE9, 0xE9, 0xF2);
        internal static readonly Color Muted = Color.FromArgb(0x9A, 0x9A, 0xAD);
        internal static readonly Color Accent = Color.FromArgb(0x3B, 0x82, 0xF6);

        public Palette(OverlayForm overlay)
        {
            this.overlay = overlay;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            BackColor = Bg;
            Text = "Screen Paint";
            AutoScaleMode = AutoScaleMode.None;
            KeyPreview = true;

            dpi = CurrentDpi();
            BuildLayout();

            icon = new MiniIcon(this);
            StartWatch();

            overlay.StrokesChanged += delegate { UpdateStatus(); };
            overlay.DrawModeChanged += delegate { UpdateModeButton(); };
            overlay.RegionCaptured += SaveRegionShot;
            UpdateModeButton();
            UpdateStatus();
        }

        private int S(int value) { return (int)Math.Round(value * dpi / 96.0); }

        private float K { get { return dpi / 96f; } }

        private static int CurrentDpi()
        {
            return Native.ScreenDpi();
        }

        private void BuildLayout()
        {
            // Ширина 320 — чтобы поместилась самая длинная подпись («marker+»,
            // «triangle»); высота выросла на ряд инструментов-фигур и на подпись.
            ClientSize = new Size(S(320), S(350));
            BuildHeader();
            BuildTools();
            BuildColors();
            BuildEffect();
            BuildSize();
            BuildActions();
            BuildStatus();
        }

        private void RebuildUi()
        {
            SuspendLayout();
            for (int i = Controls.Count - 1; i >= 0; i--)
            {
                Control c = Controls[i];
                Controls.RemoveAt(i);
                c.Dispose();
            }
            swatches.Clear();
            toolMap.Clear();
            sizeSlider = null; colorPreview = null; brushPreview = null;
            modeButton = null; statusLabel = null; effectButton = null;

            BuildLayout();

            UpdateModeButton();
            UpdateStatus();
            ResumeLayout(true);
        }

        private void ApplyDpi(int value)
        {
            if (value <= 0) value = 96;
            if (value == dpi) return;
            dpi = value;
            RebuildUi();
            ClampToScreen();
            if (icon != null)
            {
                icon.ApplyDpi(value);
                if (collapsed) icon.ClampToScreen();
                else PlaceIconAbovePanel();
            }
            SyncHole();
            Program.Log(string.Format("palette dpi={0} size={1}x{2}", dpi, Width, Height));
        }

        public void DebugSetDpi(int value) { ApplyDpi(value); }

        public void DebugResetDpi() { ApplyDpi(Native.ScreenDpi()); }

        private void ClampToScreen()
        {
            int sw = Native.GetSystemMetrics(Native.SM_CXSCREEN);
            int sh = Native.GetSystemMetrics(Native.SM_CYSCREEN);
            int x = Location.X, y = Location.Y;
            if (x + Width > sw) x = sw - Width;
            if (y + Height > sh) y = sh - Height;
            if (x < 0) x = 0;
            if (y < 0) y = 0;
            if (x != Location.X || y != Location.Y) Location = new Point(x, y);
        }

        public void PlaceTopRight(int vx, int vy)
        {
            int x = Native.GetSystemMetrics(Native.SM_CXSCREEN) - Width - S(24);
            if (x < vx) x = vx + S(8);
            Location = new Point(x, vy + S(48));
        }

        public void ShowStartup(int vx, int vy)
        {
            PlaceTopRight(vx, vy);
            icon.Show();
            PlaceIconAbovePanel();
            Show();
            SyncHole();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Native.WM_DPICHANGED)
            {
                int value = (int)(m.WParam.ToInt64() & 0xFFFF);
                base.WndProc(ref m);
                ApplyDpi(value);
                return;
            }
            if (m.Msg == Native.WM_DISPLAYCHANGE)
            {
                base.WndProc(ref m);
                ClampToScreen();
                SyncHole();
                return;
            }
            base.WndProc(ref m);
        }

        private Button FlatButton(string text, EventHandler onClick, int width, int height)
        {
            Button b = new Button();
            b.Text = text;
            b.Size = new Size(width, height);
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = Accent;
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(0x1D, 0x4E, 0xD8);
            b.BackColor = Bg2;
            b.ForeColor = Fg;
            b.Font = FitFont(text, width);
            b.Cursor = Cursors.Hand;
            b.TabStop = false;
            b.Click += onClick;
            return b;
        }

        private Font FitFont(string text, int width)
        {
            Font f = fontBtn;
            float size = fontBtn.Size;
            int limit = width - S(8);
            while (size > 7f && TextRenderer.MeasureText(text, f).Width > limit)
            {
                size -= 0.5f;
                string key = size.ToString("0.0");
                Font cached;
                if (!fonts.TryGetValue(key, out cached))
                {
                    cached = new Font("Segoe UI", size);
                    fonts.Add(key, cached);
                }
                f = cached;
            }
            return f;
        }

        private void BuildHeader()
        {
            Panel header = new Panel();
            header.Location = new Point(0, 0);
            header.Size = new Size(ClientSize.Width, S(30));
            header.BackColor = Bg2;
            Controls.Add(header);

            int btnW = S(30), btnH = S(24);

            Label title = new Label();
            title.Text = "  Screen Paint";
            title.ForeColor = Fg;
            title.BackColor = Bg2;
            title.Font = fontTitle;
            title.AutoSize = false;
            title.TextAlign = ContentAlignment.MiddleLeft;
            title.Location = new Point(0, 0);
            title.Size = new Size(Math.Max(S(80), header.Width - btnW * 2 - S(4)), header.Height);
            title.MouseDown += HeaderMouseDown;
            title.MouseMove += HeaderMouseMove;
            title.MouseUp += HeaderMouseUp;
            header.Controls.Add(title);

            Button hide = FlatButton("-", delegate { Collapse(); }, btnW, btnH);
            hide.Location = new Point(header.Width - btnW * 2 - S(2), S(3));
            hide.MouseDown += HeaderMouseDown;
            hide.MouseMove += HeaderMouseMove;
            hide.MouseUp += HeaderMouseUp;
            header.Controls.Add(hide);

            Button close = FlatButton("X", delegate { Application.Exit(); }, btnW, btnH);
            close.Location = new Point(header.Width - btnW - S(2), S(3));
            close.FlatAppearance.MouseOverBackColor = Color.FromArgb(0xDC, 0x26, 0x26);
            header.Controls.Add(close);
        }

        private void HeaderMouseDown(object sender, MouseEventArgs e) { HeaderMouseDownImpl(e); }

        private void HeaderMouseDownImpl(MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            dragging = true;
            Point screen = Cursor.Position;
            dragOffset = new Point(screen.X - Location.X, screen.Y - Location.Y);
        }

        private void HeaderMouseMove(object sender, MouseEventArgs e)
        {
            if (!dragging) return;
            Point screen = Cursor.Position;
            Location = new Point(screen.X - dragOffset.X, screen.Y - dragOffset.Y);
        }

        private void HeaderMouseUp(object sender, MouseEventArgs e) { dragging = false; }

        private void BuildTools()
        {
            const int perRow = 5;
            for (int i = 0; i < ToolKeys.Length; i++)
            {
                string key = ToolKeys[i];
                Button b = FlatButton(ToolNames[i], delegate { SetTool(key); }, S(60), S(28));
                b.Location = new Point(S(8) + (i % perRow) * S(61), S(38) + (i / perRow) * S(32));
                if (key == "text")
                    tips.SetToolTip(b, "signature: pick the colour and size, click the screen, type, enter");
                if (key == "shot")
                    tips.SetToolTip(b, "screenshot: drag a region — saves the desktop under it together with your drawings");
                Controls.Add(b);
                toolMap[key] = b;
            }
            RefreshTools();
        }

        private void BuildColors()
        {
            for (int i = 0; i < Colors.Length; i++)
            {
                Color c = Colors[i];
                Panel sw = new Panel();
                sw.Size = new Size(S(24), S(22));
                sw.BackColor = c;
                sw.Location = new Point(S(8) + (i % 7) * S(26), S(138) + (i / 7) * S(26));
                sw.Cursor = Cursors.Hand;
                sw.Click += delegate { SetColor(c); };
                sw.Paint += SwatchPaint;
                Controls.Add(sw);
                swatches.Add(sw);
            }

            Button custom = FlatButton("c", delegate { PickCustomColor(); }, S(30), S(22));
            custom.Location = new Point(S(194), S(138));
            Controls.Add(custom);

            colorPreview = new Panel();
            colorPreview.Size = new Size(S(30), S(22));
            colorPreview.Location = new Point(S(228), S(138));
            colorPreview.BackColor = overlay.BrushColor;
            Controls.Add(colorPreview);

            Label hint = new Label();
            hint.Text = "color";
            hint.ForeColor = Muted;
            hint.AutoSize = true;
            hint.Location = new Point(S(262), S(142));
            Controls.Add(hint);
        }

        private void BuildEffect()
        {
            effectButton = FlatButton(EffectNames[0], delegate { CycleEffect(); }, S(116), S(22));
            effectButton.Location = new Point(S(196), S(102));
            Controls.Add(effectButton);
            RefreshEffect();
        }

        private void CycleEffect()
        {
            int i = Array.IndexOf(EffectKeys, overlay.ShapeEffect);
            if (i < 0) i = 0;
            overlay.ShapeEffect = EffectKeys[(i + 1) % EffectKeys.Length];
            RefreshEffect();
        }

        private void RefreshEffect()
        {
            if (effectButton == null) return;
            int i = Array.IndexOf(EffectKeys, overlay.ShapeEffect);
            if (i < 0) i = 0;
            bool usable = OverlayForm.IsClosedShape(overlay.Tool);
            effectButton.Text = EffectNames[i];
            effectButton.Enabled = usable;
            effectButton.BackColor = (usable && i > 0) ? Accent : Bg2;
            effectButton.ForeColor = usable ? (i > 0 ? Color.White : Fg) : Muted;
            tips.SetToolTip(effectButton, EffectTips[i]);
        }

        private void SwatchPaint(object sender, PaintEventArgs e)
        {
            Panel p = (Panel)sender;
            Color c = p.BackColor;
            Color brush = overlay.BrushColor;
            bool selected = c.ToArgb() == brush.ToArgb();
            using (Pen pen = new Pen(selected ? Color.White : Color.FromArgb(0x44, 0x44, 0x52),
                                    selected ? 2f * K : 1f * K))
                e.Graphics.DrawRectangle(pen, 0, 0, p.Width - 1, p.Height - 1);
        }

        private void BuildSize()
        {
            Label label = new Label();
            label.Text = "size";
            label.ForeColor = Muted;
            label.AutoSize = true;
            label.Location = new Point(S(8), S(196));
            Controls.Add(label);

            sizeSlider = new Slider();
            sizeSlider.UiScale = K;
            sizeSlider.Location = new Point(S(58), S(192));
            sizeSlider.Size = new Size(S(150), S(22));
            sizeSlider.Minimum = 1;
            sizeSlider.Maximum = 60;
            sizeSlider.Value = overlay.BrushSize;
            sizeSlider.ValueChanged += delegate { SetSize(sizeSlider.Value); };
            Controls.Add(sizeSlider);

            brushPreview = new Panel();
            brushPreview.Size = new Size(S(64), S(32));
            brushPreview.Location = new Point(S(220), S(190));
            brushPreview.BackColor = Bg2;
            brushPreview.Paint += BrushPreviewPaint;
            Controls.Add(brushPreview);
        }

        private void BrushPreviewPaint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Bg2);
            float w = OverlayForm.ToolWidth(overlay.Tool, overlay.BrushSize);
            float d = Math.Max(3f, Math.Min(26f, w * 26f / 60f + 3f)) * K;
            float cx = brushPreview.Width / 2f, cy = brushPreview.Height / 2f;
            using (SolidBrush b = new SolidBrush(Color.FromArgb(
                overlay.Tool == "marker" ? 150 : (overlay.Tool == "highlighter" ? 110 : 255), overlay.BrushColor)))
                g.FillEllipse(b, cx - d / 2f, cy - d / 2f, d, d);
            if (overlay.Tool == "eraser")
            {
                using (Pen p = new Pen(Color.White, 2f * K))
                    g.DrawLine(p, cx - 7 * K, cy + 7 * K, cx + 7 * K, cy - 7 * K);
            }
        }

        private void BuildActions()
        {
            Button undo = FlatButton("undo", delegate { overlay.Undo(); }, S(96), S(28));
            undo.Location = new Point(S(8), S(228));
            Controls.Add(undo);

            Button clear = FlatButton("clear", delegate { overlay.ClearAll(); }, S(96), S(28));
            clear.Location = new Point(S(106), S(228));
            Controls.Add(clear);

            Button save = FlatButton("save PNG", delegate { SavePng(); }, S(100), S(28));
            save.Location = new Point(S(204), S(228));
            Controls.Add(save);
        }

        private void BuildStatus()
        {
            modeButton = new Button();
            modeButton.Size = new Size(ClientSize.Width - S(16), S(30));
            modeButton.Location = new Point(S(8), S(264));
            modeButton.FlatStyle = FlatStyle.Flat;
            modeButton.FlatAppearance.BorderSize = 0;
            modeButton.Font = fontMode;
            modeButton.ForeColor = Color.White;
            modeButton.Cursor = Cursors.Hand;
            modeButton.TabStop = false;
            modeButton.Click += delegate { overlay.SetDrawMode(!overlay.DrawMode); };
            Controls.Add(modeButton);

            statusLabel = new Label();
            statusLabel.Location = new Point(S(8), S(300));
            statusLabel.Size = new Size(ClientSize.Width - S(16), S(42));
            statusLabel.ForeColor = Muted;
            statusLabel.BackColor = Bg;
            statusLabel.Font = fontStatus;
            statusLabel.Text = "";
            Controls.Add(statusLabel);
        }

        private void SetTool(string key)
        {
            overlay.Tool = key;
            RefreshTools();
            brushPreview.Invalidate();
            UpdateStatus();
        }

        private void RefreshTools()
        {
            foreach (KeyValuePair<string, Button> pair in toolMap)
            {
                bool active = pair.Key == overlay.Tool;
                pair.Value.BackColor = active ? Accent : Bg2;
                pair.Value.ForeColor = active ? Color.White : Fg;
            }
            RefreshEffect();
        }

        private void SetColor(Color c)
        {
            overlay.BrushColor = c;
            colorPreview.BackColor = c;
            foreach (Panel p in swatches) p.Invalidate();
            brushPreview.Invalidate();
            UpdateStatus();
        }

        private void PickCustomColor()
        {
            using (ColorDialog dlg = new ColorDialog())
            {
                dlg.Color = overlay.BrushColor;
                dlg.FullOpen = true;
                if (ShowModal(dlg) == DialogResult.OK) SetColor(dlg.Color);
            }
        }

        private void SetSize(int size)
        {
            overlay.BrushSize = size;
            brushPreview.Invalidate();
            UpdateStatus();
        }

        public bool Collapsed { get { return collapsed; } }

        public int SinceExpand { get { return (int)(DateTime.Now - expandedAt).TotalMilliseconds; } }

        public void Collapse()
        {
            if (collapsed) return;
            collapsed = true;
            autoHide = false;
            awayTicks = 0;
            Hide();
            overlay.ClearHole("panel");
            if (icon != null) icon.Open = false;
            Program.Log("panel: collapsed to icon");
        }

        public void Expand() { Reveal(false); }

        public void ExpandFromIcon() { Reveal(true); }

        private void Reveal(bool autoHide)
        {
            if (!collapsed && Visible) return;
            this.autoHide = autoHide;
            collapsed = false;
            awayTicks = 0;
            if (icon != null) PlaceNear(icon.ButtonBounds);
            Show();
            TopMost = true;
            if (icon != null) icon.Open = true;
            SyncHole();
            expandedAt = DateTime.Now;
            Program.Log("panel: expanded");
        }

        public void FollowIcon()
        {
            if (icon == null) return;
            if (!collapsed) PlaceNear(icon.ButtonBounds);
            SyncHole();
        }

        private void TogglePanel()
        {
            if (collapsed) Expand(); else Collapse();
        }

        private void PlaceNear(Rectangle anchor)
        {
            Rectangle screen = Screen.FromRectangle(anchor).WorkingArea;
            int gap = S(4);
            int x = anchor.Right - Width;
            int y = anchor.Bottom + gap;
            if (y + Height > screen.Bottom) y = anchor.Top - Height - gap;
            if (x + Width > screen.Right) x = screen.Right - Width;
            if (x < screen.Left) x = screen.Left;
            if (y < screen.Top) y = screen.Top;
            if (y + Height > screen.Bottom) y = screen.Bottom - Height;
            Location = new Point(x, y);
        }

        private void PlaceIconAbovePanel()
        {
            if (icon == null) return;
            Rectangle screen = Screen.FromRectangle(Bounds).WorkingArea;
            int gap = S(4);
            int side = icon.ButtonPixels;
            int x = Right - side;
            int y = Top - side - gap;
            if (y < screen.Top)
            {
                if (Right + gap + side <= screen.Right) { x = Right + gap; y = Top; }
                else y = Top + gap;
            }
            if (x + side > screen.Right) x = screen.Right - side;
            if (x < screen.Left) x = screen.Left;
            if (y < screen.Top) y = screen.Top;
            icon.SetLocation(new Point(x, y));
        }

        private void StartWatch()
        {
            watch = new Timer();
            watch.Interval = 200;
            watch.Tick += delegate
            {
                if (!autoHide || collapsed || !Visible || modal > 0 || icon == null) return;
                if (Control.MouseButtons != MouseButtons.None) { awayTicks = 0; return; }
                Point p = Control.MousePosition;
                if (Bounds.Contains(p) || icon.ButtonBounds.Contains(p)) { awayTicks = 0; return; }
                awayTicks++;
                if (awayTicks >= 4) Collapse();
            };
            watch.Start();
        }

        private DialogResult ShowModal(Func<DialogResult> show)
        {
            modal++;
            try { return show(); }
            finally { modal--; awayTicks = 0; }
        }

        private DialogResult ShowModal(CommonDialog dlg)
        {
            return ShowModal(delegate { return dlg.ShowDialog(this); });
        }

        private void UpdateModeButton()
        {
            if (overlay.DrawMode)
            {
                modeButton.Text = "drawing  ·  Ctrl+Alt+D";
                modeButton.BackColor = Color.FromArgb(0xDC, 0x26, 0x26);
                modeButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(0xB9, 0x1C, 0x1C);
            }
            else
            {
                modeButton.Text = "viewing  ·  Ctrl+Alt+D";
                modeButton.BackColor = Accent;
                modeButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(0x1D, 0x4E, 0xD8);
            }
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            if (statusLabel == null) return;
            if (sizeSlider != null && sizeSlider.Value != overlay.BrushSize) sizeSlider.Value = overlay.BrushSize;
            if (brushPreview != null) brushPreview.Invalidate();
            statusLabel.Text = string.Format(
                "lines: {0}   ·   size: {1}px   ·   shift = straight\r\n" +
                "Ctrl+Alt: D mode · Z undo · X clear · S save\r\n" +
                "H panel · Q exit · text: click + type + enter",
                overlay.StrokeCount, overlay.BrushSize);
        }

        public void SavePng()
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Title = "save paint";
                dlg.Filter = "PNG (*.png)|*.png";
                dlg.FileName = "screen-paint.png";
                if (ShowModal(dlg) != DialogResult.OK) return;
                try
                {
                    overlay.SavePng(dlg.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("save error: " + ex.Message, "Screen Paint");
                }
            }
        }

        private void SaveRegionShot(Bitmap bmp)
        {
            try
            {
                using (SaveFileDialog dlg = new SaveFileDialog())
                {
                    dlg.Title = "save screenshot";
                    dlg.Filter = "PNG (*.png)|*.png";
                    dlg.FileName = "screen-shot.png";
                    if (ShowModal(dlg) != DialogResult.OK) return;
                    bmp.Save(dlg.FileName, ImageFormat.Png);
                    Program.Log("shot saved: " + dlg.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("save error: " + ex.Message, "Screen Paint");
            }
        }

        public void HandleHotkey(string action)
        {
            switch (action)
            {
                case "toggle": overlay.SetDrawMode(!overlay.DrawMode); break;
                case "undo": overlay.Undo(); break;
                case "clear": overlay.ClearAll(); break;
                case "save": SavePng(); break;
                case "panel": TogglePanel(); break;
                case "quit": Application.Exit(); break;
            }
        }

        public void SyncHole()
        {
            if (icon != null)
            {
                Rectangle b = icon.ButtonBounds;
                b.Inflate(2, 2);
                overlay.SetHole("icon", b);
            }
            if (collapsed || !Visible) { overlay.ClearHole("panel"); return; }
            overlay.SetHole("panel", Bounds);
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            SyncHole();
        }

        protected override void OnMove(EventArgs e)
        {
            base.OnMove(e);
            SyncHole();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            SyncHole();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            Native.AddExStyle(Handle, Native.WS_EX_TOOLWINDOW);
        }
    }

    internal class MiniIcon : Form
    {
        private const int LogicalSize = 30;

        private readonly Palette palette;
        private readonly Timer hover;
        private readonly Bitmap glyph;
        private int dpi = 96;
        private bool hovered;
        private bool open;
        private bool dragging;
        private bool moved;
        private Point grab;

        private IntPtr hdcScreen, hdcMem, hdcOld, hDib;
        private Bitmap surface;
        private Graphics surfaceG;
        private int button;

        public MiniIcon(Palette palette)
        {
            this.palette = palette;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            AutoScaleMode = AutoScaleMode.None;
            Text = "Screen Paint icon";
            Cursor = Cursors.Hand;
            SetStyle(ControlStyles.Opaque | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            DoubleBuffered = false;

            dpi = Native.ScreenDpi();
            button = S(LogicalSize);
            Size = new Size(button, button);
            glyph = LoadGlyph();

            hover = new Timer();
            hover.Interval = 140;
            hover.Tick += delegate { hover.Stop(); palette.ExpandFromIcon(); };
        }

        private static Bitmap LoadGlyph()
        {
            try
            {
                using (Icon ico = Icon.ExtractAssociatedIcon(Application.ExecutablePath))
                    if (ico != null) return ico.ToBitmap();
            }
            catch (Exception) { }
            return null;
        }

        private int S(int value) { return (int)Math.Round(value * dpi / 96.0); }

        public Rectangle ButtonBounds
        {
            get { return new Rectangle(Left, Top, button, button); }
        }

        public int ButtonPixels { get { return button; } }

        public bool Open { set { if (open == value) return; open = value; Redraw(); } }

        public void ApplyDpi(int value)
        {
            if (value <= 0) value = 96;
            if (value == dpi) return;
            dpi = value;
            button = S(LogicalSize);
            if (IsHandleCreated)
            {
                Size = new Size(button, button);
                EnsureSurface();
                ApplyRegion();
                Redraw();
            }
        }

        public void ClampToScreen() { SetLocation(Location); }

        public void SetLocation(Point at)
        {
            Location = Clamp(at);
            KeepOnTop();
            Push();
        }

        private void KeepOnTop()
        {
            if (!IsHandleCreated) return;
            Native.SetWindowPos(Handle, Native.HWND_TOPMOST, 0, 0, 0, 0,
                Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE);
        }

        private Point Clamp(Point p)
        {
            Rectangle v = SystemInformation.VirtualScreen;
            if (p.X > v.Right - button) p.X = v.Right - button;
            if (p.Y > v.Bottom - button) p.Y = v.Bottom - button;
            if (p.X < v.Left) p.X = v.Left;
            if (p.Y < v.Top) p.Y = v.Top;
            return p;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= Native.WS_EX_LAYERED | Native.WS_EX_TOOLWINDOW | Native.WS_EX_NOACTIVATE;
                return cp;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            EnsureSurface();
            ApplyRegion();
            Redraw();
        }

        protected override void OnPaintBackground(PaintEventArgs e) { /* рисуем сами через UpdateLayeredWindow */ }
        protected override void OnPaint(PaintEventArgs e) { }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            ReleaseSurface();
            base.OnHandleDestroyed(e);
        }

        private void EnsureSurface()
        {
            if (surfaceG != null && surface.Width == Width && surface.Height == Height) return;
            ReleaseSurface();

            hdcScreen = Native.GetDC(IntPtr.Zero);
            Native.BITMAPINFO bmi = new Native.BITMAPINFO();
            bmi.bmiHeader.biSize = Marshal.SizeOf(typeof(Native.BITMAPINFOHEADER));
            bmi.bmiHeader.biWidth = Width;
            bmi.bmiHeader.biHeight = -Height;         // top-down
            bmi.bmiHeader.biPlanes = 1;
            bmi.bmiHeader.biBitCount = 32;
            bmi.bmiHeader.biCompression = 0;          // BI_RGB
            IntPtr bits;
            hDib = Native.CreateDIBSection(hdcScreen, ref bmi, 0, out bits, IntPtr.Zero, 0);
            if (hDib == IntPtr.Zero) throw new Exception("CreateDIBSection failed");
            hdcMem = Native.CreateCompatibleDC(hdcScreen);
            hdcOld = Native.SelectObject(hdcMem, hDib);
            surface = new Bitmap(Width, Height, Width * 4, PixelFormat.Format32bppPArgb, bits);
            surfaceG = Graphics.FromImage(surface);
        }

        private void ReleaseSurface()
        {
            if (surfaceG != null) { surfaceG.Dispose(); surfaceG = null; }
            if (surface != null) { surface.Dispose(); surface = null; }
            if (hdcMem != IntPtr.Zero && hdcOld != IntPtr.Zero) Native.SelectObject(hdcMem, hdcOld);
            hdcOld = IntPtr.Zero;
            if (hDib != IntPtr.Zero) { Native.DeleteObject(hDib); hDib = IntPtr.Zero; }
            if (hdcMem != IntPtr.Zero) { Native.DeleteDC(hdcMem); hdcMem = IntPtr.Zero; }
            if (hdcScreen != IntPtr.Zero) { Native.ReleaseDC(IntPtr.Zero, hdcScreen); hdcScreen = IntPtr.Zero; }
        }

        private void ApplyRegion()
        {
            if (!IsHandleCreated) return;
            int d = Math.Max(4, S(18));
            IntPtr rgn = Native.CreateRoundRectRgn(0, 0, button + 1, button + 1, d, d);
            Native.SetWindowRgn(Handle, rgn, true);
        }

        private void Redraw()
        {
            if (surfaceG == null || !IsHandleCreated) return;
            Graphics g = surfaceG;
            g.CompositingMode = CompositingMode.SourceCopy;
            g.Clear(Color.Transparent);
            g.CompositingMode = CompositingMode.SourceOver;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            Rectangle r = new Rectangle(0, 0, button - 1, button - 1);
            using (GraphicsPath path = RoundRect(r, S(9)))
            {
                using (SolidBrush b = new SolidBrush(open || hovered ? Palette.BgHover : Palette.Bg))
                    g.FillPath(b, path);
                using (Pen p = new Pen(open || hovered ? Palette.Accent : Palette.Muted, Math.Max(1, S(1))))
                    g.DrawPath(p, path);
            }
            int inset = S(6);
            Rectangle inner = new Rectangle(inset, inset, button - inset * 2, button - inset * 2);
            if (inner.Width > 0 && inner.Height > 0)
            {
                if (glyph != null) g.DrawImage(glyph, inner);
                else using (SolidBrush dot = new SolidBrush(Palette.Accent)) g.FillEllipse(dot, inner);
            }
            Push();
        }

        private static GraphicsPath RoundRect(Rectangle r, int radius)
        {
            int d = Math.Max(2, radius * 2);
            GraphicsPath path = new GraphicsPath();
            path.AddArc(r.Left, r.Top, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void Push()
        {
            if (hdcMem == IntPtr.Zero || !IsHandleCreated) return;
            Native.POINT dst = new Native.POINT(Left, Top);
            Native.POINT src = new Native.POINT(0, 0);
            Native.SIZE size = new Native.SIZE(Width, Height);
            Native.BLENDFUNCTION bf = new Native.BLENDFUNCTION();
            bf.BlendOp = Native.AC_SRC_OVER;
            bf.BlendFlags = 0;
            bf.SourceConstantAlpha = 255;
            bf.AlphaFormat = Native.AC_SRC_ALPHA;
            Native.UpdateLayeredWindow(Handle, hdcScreen, ref dst, ref size, hdcMem, ref src, 0, ref bf, Native.ULW_ALPHA);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            hovered = true;
            Redraw();
            if (palette.Collapsed) hover.Start();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            hovered = false;
            hover.Stop();
            Redraw();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;
            dragging = true;
            moved = false;
            Point p = PointToScreen(new Point(e.X, e.Y));
            grab = new Point(p.X - Left, p.Y - Top);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!dragging) return;
            Point p = PointToScreen(new Point(e.X, e.Y));
            Point want = new Point(p.X - grab.X, p.Y - grab.Y);
            if (!moved && Math.Abs(want.X - Left) + Math.Abs(want.Y - Top) < 3) return;
            moved = true;
            SetLocation(want);
            palette.FollowIcon();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            bool wasDrag = moved;
            dragging = false;
            moved = false;
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left || wasDrag) return;
            if (palette.Collapsed) { palette.ExpandFromIcon(); return; }
            if (palette.SinceExpand > 600) palette.Collapse();
        }
    }

    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (HasArg(args, "--selftest"))
            {
                try { if (!Native.AttachConsole(Native.ATTACH_PARENT_PROCESS)) Native.AllocConsole(); }
                catch (Exception) { }
            }

            try
            {
                return Run(args);
            }
            catch (Exception ex)
            {
                Log("FATAL: " + ex);
                MessageBox.Show("Screen Paint: error\r\n\r\n" + ex, "Screen Paint");
                return 1;
            }
        }

        private static int Run(string[] args)
        {
            Native.TryEnableDpiAwareness();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += delegate(object sender, System.Threading.ThreadExceptionEventArgs e)
            {
                Log("THREAD: " + e.Exception);
            };

            bool demo = HasArg(args, "--demo");
            string selftest = ArgValue(args, "--selftest");
            double demoSeconds = ParseDouble(ArgValue(args, "--demo"), 8.0);
            int forcedDpi = (int)ParseDouble(ArgValue(args, "--dpi"), 0);
            bool noActivate = ArgValue(args, "--focus") == null;

            int vx, vy, vw, vh;
            Native.GetVirtualScreen(out vx, out vy, out vw, out vh);
            if (vw <= 0 || vh <= 0) { vx = 0; vy = 0; vw = 1920; vh = 1080; }

            OverlayForm overlay;
            try
            {
                overlay = new OverlayForm(vx, vy, vw, vh, noActivate);
            }
            catch (Exception ex)
            {
                MessageBox.Show("can't create window:\r\n" + ex, "Screen Paint");
                return 1;
            }

            Palette palette = new Palette(overlay);
            if (forcedDpi >= 96) palette.DebugSetDpi(forcedDpi);
            if (HasArg(args, "--dpi-flip"))
            {
                Timer flip = new Timer();
                flip.Interval = 2000;
                flip.Tick += delegate { flip.Stop(); palette.DebugResetDpi(); };
                flip.Start();
            }
            overlay.HotkeyPressed += delegate(object s, string action) { palette.HandleHotkey(action); };

            if (selftest != null) return SelfTest(overlay, selftest, vx, vy, vw, vh);

            if (demo)
            {
                DrawDemo(overlay, vw, vh);
                overlay.Show();
                palette.ShowStartup(vx, vy);
                overlay.RebuildCommitted();
                overlay.PresentAll();
                Timer t = new Timer();
                t.Interval = (int)(demoSeconds * 1000);
                t.Tick += delegate { t.Stop(); Application.Exit(); };
                t.Start();
                Application.Run();
                return 0;
            }

            overlay.Show();
            palette.ShowStartup(vx, vy);
            overlay.RebuildCommitted();
            overlay.PresentAll();
            Application.Run();
            return 0;
        }

        private static void DrawDemo(OverlayForm overlay, int vw, int vh)
        {
            Stroke a = new Stroke("pen", Color.FromArgb(255, 0, 0), 16, 1);
            a.Points.Add(new PointF(120, 120)); a.Points.Add(new PointF(vw - 120, 120));
            Stroke b = new Stroke("marker", Color.FromArgb(0, 200, 0), 40, 2);
            b.Points.Add(new PointF(120, 320)); b.Points.Add(new PointF(vw - 120, 340));
            Stroke c = new Stroke("highlighter", Color.FromArgb(255, 214, 10), 60, 3);
            c.Points.Add(new PointF(120, 520)); c.Points.Add(new PointF(vw - 120, 520));
            Stroke d = new Stroke("spray", Color.FromArgb(10, 132, 255), 26, 4);
            d.Points.Add(new PointF(600, 820)); d.Points.Add(new PointF(1400, 840)); d.Points.Add(new PointF(2200, 820));
            overlay.AddStroke(a); overlay.AddStroke(b); overlay.AddStroke(c); overlay.AddStroke(d);

            Stroke e = new Stroke("rect", Color.FromArgb(0, 199, 190), 10, 5);
            e.Effect = "light";
            e.Points.Add(new PointF(vw - 1000, 200)); e.Points.Add(new PointF(vw - 600, 500));
            Stroke f = new Stroke("arrow", Color.FromArgb(255, 45, 85), 12, 6);
            f.Points.Add(new PointF(vw - 1000, 700)); f.Points.Add(new PointF(vw - 600, 800));
            overlay.AddStroke(e); overlay.AddStroke(f);
        }

        internal static void Log(string text)
        {
            Console.WriteLine(text);
            Console.Out.Flush();
            try
            {
                File.AppendAllText(Path.Combine(Path.GetTempPath(), "screen-paint.log"),
                    text.Replace("\r\n", " ").Replace("\n", " ") + Environment.NewLine);
            }
            catch (Exception) { }
        }

        private static int SelfTest(OverlayForm overlay, string outDir, int vx, int vy, int vw, int vh)
        {
            if (outDir == null || outDir.Length == 0) outDir = Path.GetDirectoryName(Application.ExecutablePath);
            Log(string.Format("virtual screen: {0},{1} {2}x{3}", vx, vy, vw, vh));
            overlay.Show();
            overlay.SetToolbarHole(new Rectangle(vx + 20, vy + 20, 304, 350));

            Stroke a = new Stroke("pen", Color.FromArgb(255, 59, 48), 14, 11);
            a.Points.Add(new PointF(150, 150)); a.Points.Add(new PointF(700, 420)); a.Points.Add(new PointF(1250, 240));
            Stroke b = new Stroke("marker", Color.FromArgb(52, 199, 89), 34, 12);
            b.Points.Add(new PointF(150, 620)); b.Points.Add(new PointF(1400, 660));
            Stroke c = new Stroke("highlighter", Color.FromArgb(255, 214, 10), 60, 13);
            c.Points.Add(new PointF(150, 760)); c.Points.Add(new PointF(1400, 760));
            Stroke d = new Stroke("spray", Color.FromArgb(10, 132, 255), 30, 12345);
            d.Points.Add(new PointF(900, 900)); d.Points.Add(new PointF(1100, 930));
            Stroke e = new Stroke("pen", Color.FromArgb(255, 255, 255), 8, 15);
            e.Points.Add(new PointF(150, 1000)); e.Points.Add(new PointF(1300, 1150));
            overlay.AddStroke(a); overlay.AddStroke(b); overlay.AddStroke(c);
            overlay.AddStroke(d); overlay.AddStroke(e);

            Stroke r = new Stroke("rect", Color.FromArgb(0, 199, 190), 8, 21);
            r.Points.Add(new PointF(1700, 200)); r.Points.Add(new PointF(2100, 500));
            Stroke el = new Stroke("ellipse", Color.FromArgb(191, 90, 242), 8, 22);
            el.Points.Add(new PointF(1700, 700)); el.Points.Add(new PointF(2100, 1100));
            Stroke tr = new Stroke("triangle", Color.FromArgb(255, 214, 10), 8, 23);
            tr.Points.Add(new PointF(2300, 700)); tr.Points.Add(new PointF(2700, 1100));
            Stroke ln = new Stroke("line", Color.FromArgb(255, 255, 255), 6, 24);
            ln.Points.Add(new PointF(1700, 1300)); ln.Points.Add(new PointF(2100, 1500));
            Stroke ar = new Stroke("arrow", Color.FromArgb(255, 45, 85), 10, 25);
            ar.Points.Add(new PointF(2300, 1300)); ar.Points.Add(new PointF(2700, 1400));
            overlay.AddStroke(r); overlay.AddStroke(el); overlay.AddStroke(tr);
            overlay.AddStroke(ln); overlay.AddStroke(ar);

            overlay.RebuildCommitted();
            overlay.PresentAll();

            string file = Path.Combine(outDir, "screen-paint-selftest.png");
            overlay.SavePng(file);
            FileInfo fi = new FileInfo(file);
            Log(string.Format("png: {0} ({1} bytes)", file, fi.Length));

            using (Bitmap chk = new Bitmap(file))
            {
                Log(string.Format("png size: {0}x{1}", chk.Width, chk.Height));
                Log(string.Format("alpha on stroke: {0} (wait 255), alpha on empty: {1} (wait 0)",
                    chk.GetPixel(150, 150).A, chk.GetPixel(5, 5).A));

                int[,] expect = {
                    { 1900, 200, 255 },
                    { 1900, 350, 0 },
                    { 1900, 700, 255 },
                    { 1900, 900, 0 },
                    { 2500, 700, 255 },
                    { 2500, 950, 0 },
                    { 1900, 1400, 255 },
                    { 2500, 1350, 255 }
                };
                bool shapesOk = true;
                for (int i = 0; i < expect.GetLength(0); i++)
                {
                    int got = chk.GetPixel(expect[i, 0], expect[i, 1]).A;
                    if (got != expect[i, 2]) shapesOk = false;
                    Log(string.Format("shape pixel {0},{1}: alpha {2} (wait {3})",
                        expect[i, 0], expect[i, 1], got, expect[i, 2]));
                }
                Log("shapes: " + (shapesOk ? "OK" : "FAIL"));
            }
            Log(string.Format("layered surface alpha: on stroke {0} (wait 255), empty {1} (wait 0)",
                overlay.SurfaceAlphaAt(150, 150), overlay.SurfaceAlphaAt(5, 5)));

            Stroke tx = new Stroke("text", Color.FromArgb(255, 255, 255), 12, 26);
            tx.Text = "signature 42";
            tx.Points.Add(new PointF(300, 2100));
            overlay.AddStroke(tx);
            overlay.RebuildCommitted();
            overlay.PresentAll();
            string fileText = Path.Combine(outDir, "screen-paint-selftest-text.png");
            overlay.SavePng(fileText);
            int ink, above;
            using (Bitmap chk = new Bitmap(fileText))
            {
                ink = CountAlpha(chk, 300, 2090, 480, 100);
                above = CountAlpha(chk, 300, 2000, 480, 70);
            }
            bool textOk = ink > 400 && above == 0;
            Log(string.Format("signature: font {0}px, painted pixels {1} (wait >400), empty above {2} (wait 0)",
                OverlayForm.TextPixelSize(12), ink, above));
            Log("signature: " + (textOk ? "OK" : "FAIL"));

            bool shotOpaque = false, shotInk = false;
            Color shotPixel = Color.Empty;
            using (Bitmap shot = overlay.CaptureRegion(new RectangleF(415, 275, 120, 120)))
            {
                if (shot != null)
                {
                    shotOpaque = true;
                    for (int y = 0; y < shot.Height && shotOpaque; y++)
                        for (int x = 0; x < shot.Width; x++)
                            if (shot.GetPixel(x, y).A != 255) { shotOpaque = false; break; }
                    shotPixel = shot.GetPixel(10, 10);
                    shotInk = shotPixel.R > 200 && shotPixel.G < 90 && shotPixel.B < 90;
                }
            }
            Log(string.Format("shot: background opaque {0} (wait True), pixel on red stroke {1},{2},{3} (wait 255,59,48)",
                shotOpaque, shotPixel.R, shotPixel.G, shotPixel.B));

            Stroke gray = new Stroke("pen", Color.FromArgb(128, 128, 128), 80, 41);
            gray.Points.Add(new PointF(300, 1750)); gray.Points.Add(new PointF(1100, 1750));
            Stroke litShot = new Stroke("rect", Color.FromArgb(255, 255, 255), 6, 42);
            litShot.Effect = "light";
            litShot.Points.Add(new PointF(400, 1600)); litShot.Points.Add(new PointF(900, 1900));
            overlay.AddStroke(gray); overlay.AddStroke(litShot);
            overlay.RebuildCommitted();
            overlay.PresentAll();

            RectangleF strip = new RectangleF(300, 1650, 800, 200);
            Bitmap shotLit = overlay.CaptureRegion(strip);
            overlay.Undo();
            Bitmap shotPlain = overlay.CaptureRegion(strip);
            overlay.Undo();
            int inLit = 0, inPlain = 0, outLit = 0, outPlain = 0;
            if (shotLit != null && shotPlain != null)
            {
                inLit = Brightness(shotLit, 600 - 300, 1750 - 1650);
                inPlain = Brightness(shotPlain, 600 - 300, 1750 - 1650);
                outLit = Brightness(shotLit, 350 - 300, 1750 - 1650);
                outPlain = Brightness(shotPlain, 350 - 300, 1750 - 1650);
            }
            using (shotLit) using (shotPlain) { }
            bool shotEffect = inLit > inPlain + 20 && Math.Abs(outLit - outPlain) <= 6;
            Log(string.Format("shot effect: inside {0} vs {1} (wait brighter), outside {2} vs {3} (wait same)",
                inLit, inPlain, outLit, outPlain));
            Log("shot: " + ((shotOpaque && shotInk && shotEffect) ? "OK" : "FAIL"));

            bool effectOk = true;

            Stroke lit = new Stroke("rect", Color.FromArgb(255, 255, 255), 6, 31);
            lit.Effect = "light";
            lit.Points.Add(new PointF(500, 1700)); lit.Points.Add(new PointF(1100, 2000));
            overlay.AddStroke(lit);
            overlay.RebuildCommitted();
            overlay.PresentAll();
            string fileLit = Path.Combine(outDir, "screen-paint-selftest-light.png");
            overlay.SavePng(fileLit);
            using (Bitmap fx = new Bitmap(fileLit))
            {
                int inside = fx.GetPixel(800, 1850).A;
                int outside = fx.GetPixel(3500, 1500).A;
                effectOk = effectOk && inside == OverlayForm.LightAlpha && outside == 0;
                Log(string.Format("highlight: alpha inside {0} (wait {1}), outside {2} (wait 0)",
                    inside, OverlayForm.LightAlpha, outside));
            }

            Stroke dim = new Stroke("ellipse", Color.FromArgb(255, 255, 255), 6, 32);
            dim.Effect = "dim";
            dim.Points.Add(new PointF(1400, 1700)); dim.Points.Add(new PointF(2200, 2300));
            overlay.AddStroke(dim);
            overlay.RebuildCommitted();
            overlay.PresentAll();
            string fileDim = Path.Combine(outDir, "screen-paint-selftest-dim.png");
            overlay.SavePng(fileDim);
            using (Bitmap fx = new Bitmap(fileDim))
            {
                int hole = fx.GetPixel(1800, 2000).A;
                Color shadow = fx.GetPixel(3500, 2200);
                Color under = fx.GetPixel(150, 150);
                effectOk = effectOk && hole == 0 && shadow.A == OverlayForm.DimAlpha &&
                           under.A == 255 && under.R < 200;
                Log(string.Format(
                    "dim: alpha in selection {0} (wait 0), outside {1} (wait {2}), dimmed stroke R={3} A={4} (wait R<200, A=255)",
                    hole, shadow.A, OverlayForm.DimAlpha, under.R, under.A));
            }
            Log("effects: " + (effectOk ? "OK" : "FAIL"));

            int hw = vw / 2, hh = vh / 2;
            bool shrunk = overlay.DebugSetScreenBounds(vx, vy, hw, hh);
            Log(string.Format("resize {0}x{1}: applied={2}, alpha on scaled stroke {3} (wait 255)",
                hw, hh, shrunk, overlay.SurfaceAlphaAt(75, 75)));
            string small = Path.Combine(outDir, "screen-paint-selftest-small.png");
            overlay.SavePng(small);
            using (Bitmap chk2 = new Bitmap(small))
                Log(string.Format("small png size: {0}x{1} (wait {2}x{3})", chk2.Width, chk2.Height, hw, hh));

            bool restored = overlay.DebugSetScreenBounds(vx, vy, vw, vh);
            Log(string.Format("resize back: applied={0}, alpha on stroke {1} (wait 255)",
                restored, overlay.SurfaceAlphaAt(150, 150)));

            Log("SELFTEST OK");
            return 0;
        }

        private static bool HasArg(string[] args, string name)
        {
            for (int i = 0; i < args.Length; i++)
                if (args[i] == name || args[i].StartsWith(name + "=")) return true;
            return false;
        }

        private static string ArgValue(string[] args, string name)
        {
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == name) return (i + 1 < args.Length && !args[i + 1].StartsWith("--")) ? args[i + 1] : "";
                if (args[i].StartsWith(name + "=")) return args[i].Substring(name.Length + 1);
            }
            return null;
        }

        private static int CountAlpha(Bitmap bmp, int x0, int y0, int w, int h)
        {
            int n = 0;
            for (int y = y0; y < y0 + h && y < bmp.Height; y++)
            {
                if (y < 0) continue;
                for (int x = x0; x < x0 + w && x < bmp.Width; x++)
                {
                    if (x < 0) continue;
                    if (bmp.GetPixel(x, y).A > 0) n++;
                }
            }
            return n;
        }

        private static int Brightness(Bitmap bmp, int x, int y)
        {
            int sum = 0, n = 0;
            for (int j = y - 4; j <= y + 4; j++)
            {
                if (j < 0 || j >= bmp.Height) continue;
                for (int i = x - 4; i <= x + 4; i++)
                {
                    if (i < 0 || i >= bmp.Width) continue;
                    Color c = bmp.GetPixel(i, j);
                    sum += (c.R + c.G + c.B) / 3;
                    n++;
                }
            }
            return n == 0 ? 0 : sum / n;
        }

        private static double ParseDouble(string text, double fallback)
        {
            double value;
            if (text != null && double.TryParse(text, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out value)) return value;
            return fallback;
        }
    }
}
