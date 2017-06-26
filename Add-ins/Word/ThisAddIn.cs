using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Sockets;
using OpenM;

namespace WordAddIn2
{
    

    public partial class ThisAddIn
    {


        Socket sck;
        EndPoint epLocal, epRemote;
        int AppId = 0;

        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            if (AppId == 0) AppId = HKEY.GetRegistryValue("AppID");
            SendingParams p = new SendingParams() { actionType = 0, appType = 1, docID = AppId };

            MPV.Sending(p);
        }

        void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            if (AppId == 0) AppId = HKEY.GetRegistryValue("AppID");
            SendingParams p = new SendingParams() { actionType = 1, appType = 1, docID = AppId };
            MPV.Sending(p);
        }

        void Application_DocumentBeforePrint(Word.Document Doc, ref bool Cancel)
        {
            if (AppId == 0) AppId = HKEY.GetRegistryValue("AppID");
            SendingParams p = new SendingParams() { actionType = 2, appType = 1, docID = AppId };
            MPV.Sending(p);
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //  sck = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
            //  sck.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            //  MPV.connect(1,1);
            
           

            this.Application.DocumentBeforeSave  += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
            this.Application.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
            this.Application.DocumentBeforePrint += new Word.ApplicationEvents4_DocumentBeforePrintEventHandler(Application_DocumentBeforePrint);
            sck = MPV.Connect(1, 1);
            epLocal = MPV.getPoint(0);
            epRemote = MPV.getPoint(1);
            
            
            // byte[] buffer = new byte[1500];
            //    sck.BeginReceiveFrom(buffer, 0, buffer.Length, SocketFlags.None, ref epRemote, new AsyncCallback(MessageCallBack), buffer);
         //   AppId = 
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
