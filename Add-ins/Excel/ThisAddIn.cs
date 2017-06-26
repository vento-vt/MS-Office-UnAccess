using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using OpenM;

namespace ExcelAddIn2
{
        

    public partial class ThisAddIn
    {


        Socket sck;
        EndPoint epLocal, epRemote;
        int AppId = 0;   
     

        void app_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)

        {
            // string curTimeLong = DateTime.Now.ToLongTimeString();
            //  string currentWorkbookFileName = this.Application.ActiveWorkbook.Name;
            // string text = "\n Пользователь сохранил файл: " +
            //              currentWorkbookFileName + " в " + curTimeLong;
            //  string text = "0,0";

            //  System.IO.File.AppendAllText(@"E:\LogExcelSave.log", text);
            //Sending(text);
            if (AppId == 0) AppId = HKEY.GetRegistryValue("AppID");
            SendingParams p = new SendingParams() { actionType = 0, appType = 0, docID = AppId };
            MPV.Sending(p);

        }

        void app_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            //  string curTimeLong = DateTime.Now.ToLongTimeString();
            //  string currentWorkbookFileName = this.Application.ActiveWorkbook.Name;
            //  string text = "\n Пользователь закрыл файл: " +
            //            currentWorkbookFileName + " в " + curTimeLong;
            //   string text = "1,0";
            //  System.IO.File.AppendAllText(@"E:\LogExcelClose.log", text);
            if (AppId == 0) AppId = HKEY.GetRegistryValue("AppID");
            SendingParams p = new SendingParams() { actionType = 1, appType = 0, docID = AppId };
            //   MessageBox.Show(p.actionType.ToString() + p.appType.ToString() + p.docID.ToString());
            MPV.Sending(p);
        }

        void app_WorkbookBeforePrint(Excel.Workbook Wb, ref bool Cancel)
        {
            // string curTimeLong = DateTime.Now.ToLongTimeString();
            // string currentWorkbookFileName = this.Application.ActiveWorkbook.Name;
            //  string text = "\n Пользователь распечатал файл: " +
            //            currentWorkbookFileName + " в " + curTimeLong;
            //  string text = "2,0";
            // System.IO.File.AppendAllText(@"E:\LogExcelPrint.log", text);
            if (AppId == 0) AppId = HKEY.GetRegistryValue("AppID");
            SendingParams p = new SendingParams() { actionType = 2, appType = 0, docID = AppId };
            MPV.Sending(p);
        }


        
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            
            this.Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(app_WorkbookBeforeSave);
            this.Application.WorkbookBeforePrint += new Excel.AppEvents_WorkbookBeforePrintEventHandler(app_WorkbookBeforePrint);
            this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(app_WorkbookBeforeClose);

            sck = MPV.Connect(1, 1);
            epLocal = MPV.getPoint(0);
            epRemote = MPV.getPoint(1);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
