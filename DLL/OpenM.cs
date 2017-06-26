using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Net;
using System.Net.Sockets;

namespace OpenM
{
    #region [  Open  ]

    public class OpenExcel : MPV
    {

        public static int Open()
        {
            int type = 0;
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Workbooks.Add();
            excelApp.Visible = true;
            Process pExcel = CheckReg.getExcelProcess(excelApp.Caption);
            if (HKEY.checkMachineType(type) == false)
            {
                excelApp.Visible = false;
                excelApp.Quit();

                return 0;
            }
            return pExcel.Id;
        }

        public static int Open(string fn)
        {
            int type = 0;
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;
            excelApp.Workbooks.Open(fn);
            Process pExcel = CheckReg.getExcelProcess(excelApp.Caption);
            if (HKEY.checkMachineType(type) == false)
            {
                excelApp.Visible = false;
                excelApp.Quit();

                return 0;
            }
            return pExcel.Id;
        }



    }


    public class OpenWord : MPV
    {
        public static int Open()
        {
            int type = 1;
            var wordApp = new Word.Application();
            wordApp.Visible = true;
            Process pWord = CheckReg.getWordProcess(wordApp.Caption);
            wordApp.Documents.Add();
            if (HKEY.checkMachineType(type) == false)
            {
                wordApp.Visible = false;
                wordApp.Quit();
                return 0;
            }
            return pWord.Id;
        }

        public static int Open(string fn)
        {
            int type = 1;
            var wordApp = new Word.Application();
            wordApp.Visible = true;
            Process pWord = CheckReg.getWordProcess(wordApp.Caption);
            wordApp.Documents.Open(fn);
            if (HKEY.checkMachineType(type) == false)
            {
                wordApp.Visible = false;
                wordApp.Quit();
                return 0;
            }
            return pWord.Id;
        }

    }

    #endregion

    public struct procData
    {
        public int procId;
        public byte type;
    }


    

    
}

