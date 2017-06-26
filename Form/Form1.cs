using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OpenM;
using System.Net;
using System.Net.Sockets;

namespace WindowsFormsApplication2
{

    public partial class Form1 : Form
    {
        private IList<procData> pids;
        Socket sck;


        public void MessageEvent()
        {

        }

        public Form1()
        {
            
            InitializeComponent();
            timer2.Interval = 1000; //интервал между срабатываниями 1000 миллисекунд
            timer2.Tick += new EventHandler(timer2_Tick); //подписываемся на события Tick
            timer2.Start();  //событие запускает функцию проверки регистра надстроек
            pids = new List<procData>();
            sck = MPV.Connect(0,0);
            OpenExcel.DocEvent += OpenExcel_DocEvent;
            
        }

        private void OpenExcel_DocEvent(int arg1, int arg2,int id)
        {
            string s;
            // событие типа arg1 с ID = arg2
            if (arg2 == 0)
            {
                switch (arg1)

                {
                    case 0:
                        s = "Сохранена книга в приложении с идентификатором : " + id.ToString() + "\n";
                        
                        LogEvent(s);
                        break;
                    case 1:
                        s = "Закрыта книга в приложении с идентификатором : " + id.ToString() + "\n";
                        LogEvent(s);
                        break;
                    case 2:
                        s = "Книга распечатана в приложении с идентификатором : " + id.ToString() + "\n";
                        LogEvent(s);
                        break;
                }
            }
            if (arg2 == 1)
            {
                switch (arg1)

                {
                    case 0:
                        s = "Сохранен документ в приложении с идентификатором : " + id.ToString() + "\n";

                        LogEvent(s);
                        break;
                    case 1:
                        s = "Закрыт документ в приложении с идентификатором : " + id.ToString() + "\n";
                        LogEvent(s);
                        break;
                    case 2:
                        s = "Документ распечатан в приложении с идентификатором : " + id.ToString() + "\n";
                        LogEvent(s);
                        break;
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {

            string checkResult = CheckReg.Check(ref pids);
            if (!String.IsNullOrEmpty(checkResult))
                LogEvent(checkResult);

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (var pData in pids)
                if (pData.type == 1)
                {
                    Process.GetProcessById(pData.procId).Kill();
                }

            foreach (var pData in pids)
                if (pData.type == 0)
                {
                    Process.GetProcessById(pData.procId).Kill();
                }

            MessageBox.Show("Все процессы были закрыты, так как программа контроля закрылась");
            Dispose();

        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            int pid = OpenExcel.Open();
            if (pid != 0)
            {
                pids.Add(new procData { procId = pid, type = 0 });
                LogEvent(string.Format("Открыт процесс Excel c id={0}{1}", pid, Environment.NewLine));
                
                HKEY.AddReg(pid);
            }
        }

        private void btnOpenWord_Click(object sender, EventArgs e)
        {

            int pid = OpenWord.Open();
            if (pid != 0)
            {
                pids.Add(new procData { procId = pid, type = 1 });
                LogEvent(string.Format("Открыт процесс Word c id={0}{1}", pid, Environment.NewLine));
                //  System.Threading.Thread.Sleep(5000);
                //   MPV.Sending(pid);
                HKEY.AddReg(pid);
            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            // Displays an OpenFileDialog so the user can select a dok.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Word Documents|*.docx|Excel Files|*.xlsx";
            openFileDialog1.Title = "Select a document";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .doc .xls file was selected, open it.

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string ext = Path.GetExtension(openFileDialog1.FileName);
                string fn = openFileDialog1.FileName;

                if (ext == ".docx" || ext == ".doc")
                {
                    int pid = OpenWord.Open(fn);
                    if (pid != 0)
                    {
                        pids.Add(new procData { procId = pid, type = 1 });
                        LogEvent(string.Format("Открыт процесс Word c id={0}{1}", pid, Environment.NewLine));
                    }

                }
                if (ext == ".xls" || ext == ".xlsx")
                {
                    int pid = OpenExcel.Open(fn);
                    if (pid != 0)
                    {
                        pids.Add(new procData { procId = pid, type = 0 });
                        LogEvent(string.Format("Открыт процесс Excel c id={0}{1}", pid, Environment.NewLine));
                    }

                }
            }
        }

        private void LogEvent(string eventText)
        {

            tbLog.Text += eventText;

         
        }


    }
}

