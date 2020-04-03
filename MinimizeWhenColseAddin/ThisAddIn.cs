using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace MinimizeWhenColseAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var currentObject = Globals.ThisAddIn.Application.ActiveExplorer();
            var winh = currentObject as IOleWindow;
            IntPtr win;
            winh.GetWindow(out win);
            var nativeWindow = new CloseHandler(win);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //注: Outlook はこのイベントを発行しなくなりました。Outlook が
            //    を Outlook のシャットダウン時に実行する必要があります。https://go.microsoft.com/fwlink/?LinkId=506785 をご覧ください
            //Application.Startup -= OutlookStartup;
        }

        private void OutlookStartup()
        {
            Application.ActiveExplorer().WindowState = Outlook.OlWindowState.olMinimized;
        }

        private void OutlookClose()
        {
            // Restart outlook minimized
            ProcessStartInfo psiOutlook = new ProcessStartInfo("OUTLOOK.EXE", "/recycle")
            {
                WindowStyle = ProcessWindowStyle.Minimized
            };
            Process.Start(psiOutlook);
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
