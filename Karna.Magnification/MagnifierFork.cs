using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;

namespace Karna.Magnification.Fork
{
    public class Magnifier : IDisposable
    {
        private readonly Form form;
        private IntPtr hwndMag;
        private float magnification;
        private readonly bool initialized;
        private RECT magWindowRect = new RECT();
        private Timer timer;

        private static NativeMethods.ImageScalingCallback cbImg = new NativeMethods.ImageScalingCallback(Callback);

        public Magnifier(Form form)
        {
            magnification = 1.0f;
            this.form = form ?? throw new ArgumentNullException(nameof(form));
            form.Resize += new EventHandler(Form_Resize);
            form.FormClosing += new FormClosingEventHandler(Form_FormClosing);

            timer = new Timer();
            timer.Tick += new EventHandler(Timer_Tick);

            initialized = NativeMethods.MagInitialize();
            if (initialized)
            {
                SetupMagnifier();
                timer.Interval = NativeMethods.USER_TIMER_MINIMUM;
                timer.Enabled = true;
            }
        }

        void Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer.Enabled = false;
        }

        void Timer_Tick(object sender, EventArgs e)
        {
            UpdateMaginifier();
        }

        void Form_Resize(object sender, EventArgs e)
        {
            ResizeMagnifier();
        }

        protected virtual void ResizeMagnifier()
        {
            if (initialized && (hwndMag != IntPtr.Zero))
            {
                NativeMethods.GetClientRect(form.Handle, ref magWindowRect);
                // Resize the control to fill the window.
                NativeMethods.SetWindowPos(hwndMag, IntPtr.Zero,
                    magWindowRect.left, magWindowRect.top, magWindowRect.right, magWindowRect.bottom, 0);
            }
        }

        public virtual void UpdateMaginifier()
        {
            if ((!initialized) || (hwndMag == IntPtr.Zero))
            {
                return;
            }

            bool res = NativeMethods.MagSetWindowFilterList(hwndMag, MW_FILTERMODE_EXCLUDE, 1, new[] { form.Handle });

            int width = (int)((magWindowRect.right - magWindowRect.left) / magnification);
            int height = (int)((magWindowRect.bottom - magWindowRect.top) / magnification);

            RECT sourceRect = new RECT
            {
                left = form.Bounds.Left + 10,
                top = form.Bounds.Top + 33
            };


            // Don't scroll outside desktop area.
            //if (sourceRect.left < 0)
            //{
            //    sourceRect.left = 0;
            //}
            //if (sourceRect.left > NativeMethods.GetSystemMetrics(NativeMethods.SM_CXSCREEN) - width)
            //{
            //    sourceRect.left = NativeMethods.GetSystemMetrics(NativeMethods.SM_CXSCREEN) - width;
            //}
            //sourceRect.right = sourceRect.left + width;

            //if (sourceRect.top < 0)
            //{
            //    sourceRect.top = 0;
            //}
            //if (sourceRect.top > NativeMethods.GetSystemMetrics(NativeMethods.SM_CYSCREEN) - height)
            //{
            //    sourceRect.top = NativeMethods.GetSystemMetrics(NativeMethods.SM_CYSCREEN) - height;
            //}
            sourceRect.bottom = sourceRect.top + height;

            if (form.IsDisposed)
            {
                timer.Enabled = false;
                return;
            }

            bCallbackDone = false;
            // Set the source rectangle for the magnifier control.
            NativeMethods.MagSetWindowSource(hwndMag, sourceRect);

            while (!bCallbackDone) ; 

            // Reclaim topmost status, to prevent unmagnified menus from remaining in view. 
            //NativeMethods.SetWindowPos(form.Handle, NativeMethods.HWND_TOPMOST, 0, 0, 0, 0,
            //    (int)(SetWindowPosFlags.SWP_NOACTIVATE | SetWindowPosFlags.SWP_NOMOVE | SetWindowPosFlags.SWP_NOSIZE));

            // Force redraw.
            NativeMethods.InvalidateRect(hwndMag, IntPtr.Zero, true);
        }

        public float Magnification
        {
            get => magnification;
            set
            {
                if (magnification != value)
                {
                    magnification = value;
                    // Set the magnification factor.
                    Transformation matrix = new Transformation(magnification);
                    NativeMethods.MagSetWindowTransform(hwndMag, ref matrix);
                }
            }
        }

        protected void SetupMagnifier()
        {
            if (!initialized)
            {
                return;
            }

            IntPtr hInst = NativeMethods.GetModuleHandle(null);

            // Make the window opaque.
            form.AllowTransparency = true;
            form.TransparencyKey = Color.Empty;
            form.Opacity = 255;

            // Create a magnifier control that fills the client area.
            NativeMethods.GetClientRect(form.Handle, ref magWindowRect);
            hwndMag = NativeMethods.CreateWindow((int)ExtendedWindowStyles.WS_EX_CLIENTEDGE, NativeMethods.WC_MAGNIFIER,
                "MagnifierWindow", (int)WindowStyles.WS_CHILD | (int)MagnifierStyle.MS_SHOWMAGNIFIEDCURSOR |
                (int)WindowStyles.WS_VISIBLE,
                magWindowRect.left, magWindowRect.top, magWindowRect.right, magWindowRect.bottom, form.Handle, IntPtr.Zero, hInst, IntPtr.Zero);

            if (hwndMag == IntPtr.Zero)
            {
                return;
            }

            IntPtr ptr = IntPtr.Zero;
            bool res = NativeMethods.MagSetWindowFilterList(hwndMag, MW_FILTERMODE_EXCLUDE, 1, new[] { form.Handle });
            int r = NativeMethods.MagGetWindowFilterList(hwndMag, IntPtr.Zero, 1, ref ptr);
            bool result = NativeMethods.MagSetImageScalingCallback(hwndMag, cbImg);

            // Set the magnification factor.
            Transformation matrix = new Transformation(magnification);
            NativeMethods.MagSetWindowTransform(hwndMag, ref matrix);
        }

        public const int MW_FILTERMODE_EXCLUDE = 0;
        const int BI_RGB = 0;
        public const int CBM_INIT = 0x4;
        public const int DIB_RGB_COLORS = 0;

        private static bool bCallbackDone = false;

        public static Action<Bitmap> DoStuff { get; set; }

        private static bool Callback(IntPtr hwnd, IntPtr srcdata, MAGIMAGEHEADER srcheader, ref IntPtr destdata, MAGIMAGEHEADER destheader, RECT unclipped, RECT clipped, IntPtr dirty)
        {
            if (DoStuff == null)
            {
                bCallbackDone = true;
                return true;
            }

            BITMAPINFO lpbmi = new BITMAPINFO();
            lpbmi.bmiHeader.biSize = Marshal.SizeOf(lpbmi.bmiHeader);
            lpbmi.bmiHeader.biHeight = -srcheader.height;
            lpbmi.bmiHeader.biWidth = srcheader.width;
            lpbmi.bmiHeader.biSizeImage = srcheader.cbSize;
            lpbmi.bmiHeader.biPlanes = 1;
            lpbmi.bmiHeader.biBitCount = 32;
            lpbmi.bmiHeader.biCompression = BI_RGB;
            IntPtr hDC = NativeMethods.GetWindowDC(hwnd);
            IntPtr hBitmap = NativeMethods.CreateDIBitmap(hDC, ref lpbmi.bmiHeader, CBM_INIT, srcdata, ref lpbmi, DIB_RGB_COLORS);

            try
            {
                using (var img = Image.FromHbitmap(hBitmap))
                {
                    DoStuff(img);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine($"Last Win32 Error: {Marshal.GetLastWin32Error()}{Environment.NewLine}");
            }
            finally
            {
                bCallbackDone = true;
                NativeMethods.DeleteObject(hBitmap);
                NativeMethods.DeleteDC(hDC);
            }

            return true;
        }

        protected void RemoveMagnifier()
        {
            if (initialized)
            {
                NativeMethods.MagUninitialize();
            }
        }

        #region IDisposable Members

        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (disposedValue)
            {
                return;
            }

            timer.Enabled = false;
            if (disposing)
            {
                // dispose managed state (managed objects)
                timer.Dispose();
            }

            form.Resize -= Form_Resize;

            // free unmanaged resources (unmanaged objects) and override finalizer
            RemoveMagnifier();

            // set large fields to null
            timer = null;

            disposedValue = true;
        }

        ~Magnifier()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
