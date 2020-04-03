using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MinimizeWhenColseAddin
{

    // NativeWindow class to listen to operating system messages.
    [System.Security.Permissions.PermissionSet(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public class CloseHandler : NativeWindow
    {
        #region WM_SYSCOMMAND message

        // Closes the window.
        private const int WM_SYSCOMMAND = 0x0112;

        // Closes the window.
        private const int WM_CLOSE = 0x0010;

        // Minimizes the window.
        private const int SC_MINIMIZE = 0xF020;

        #endregion

        [DllImport("user32.dll")]
        // Get the information(System message) about the specified window
        private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        // set the information(System message) about the specified window
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);



        public static void HwndHander(IntPtr hwnd)
        {
            var outlookWndProc = GetWindowLongPtr(hwnd, (int)GWL.GWL_WNDPROC);
        }

        public CloseHandler(IntPtr hwnd)
        {
            // Assigns a handle to this window.
            AssignHandle(hwnd);
        }

        /// <summary>
        /// Window was destroyed, release hook.
        /// </summary>
        public void Release()
        {
            ReleaseHandle();
        }


        /// <summary>
        /// Listen for operating system messages
        /// </summary>
        /// <param name="m"></param>
        [System.Security.Permissions.PermissionSet(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLOSE)
            {
                m.Msg = WM_SYSCOMMAND;
                m.WParam = (IntPtr)SC_MINIMIZE;
                m.LParam = IntPtr.Zero;
            }

            base.WndProc(ref m);
        }
    }
}
