using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OpenM
{
    public class CheckReg
    {
        [DllImport("user32.dll")]

        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);


        public static Process getWordProcess(string pCaption)
        {
            Process[] pWords = Process.GetProcessesByName("WINWORD");
            foreach (Process pWord in pWords)
            {
                if (pWord.MainWindowTitle == pCaption)
                { return pWord; }
            }
            return null;
        }

        public static Process getExcelProcess(string pCaption)
        {
            Process[] pExcels = Process.GetProcessesByName("EXCEL");
            foreach (Process pExcel in pExcels)
            {
                if (pExcel.MainWindowTitle == pCaption)
                { return pExcel; }
            }
            return null;
        }

        public static string Check(ref IList<procData> pids)
        {
            string r = String.Empty;
            IList<procData> pids_return = new List<procData>();

            if (HKEY.checkMachineType(1) == false)
            {
                foreach (var pData in pids)
                    if (pData.type == 1)
                    {
                        if (Process.GetProcesses().Any(p => p.Id == pData.procId))
                        {
                            Process p = Process.GetProcessById(pData.procId);
                            p.Kill();
                            r += string.Format("Процесс Word {0} был завершен системой{1}", pData.procId, Environment.NewLine);
                        }
                        else
                            r += string.Format("Процесс Word {0} был завершен пользователем{1}", pData.procId, Environment.NewLine);
                    }


            }
            else
            {
                foreach (var pData in pids)
                    if (pData.type == 1)
                    {
                        if (Process.GetProcesses().Any(p => p.Id == pData.procId))
                            pids_return.Add(pData);
                        else
                            r += string.Format("Процесс Word {0} был завершен пользователем{1}", pData.procId, Environment.NewLine);
                    }
            }

            if (HKEY.checkMachineType(0) == false)
            {
                foreach (var pData in pids)
                    if (pData.type == 0)
                    {
                        if (Process.GetProcesses().Any(p => p.Id == pData.procId))
                        {
                            Process p = Process.GetProcessById(pData.procId);
                            p.Kill();
                            r += string.Format("Процесс Excel {0} был завершен системой{1}", pData.procId, Environment.NewLine);
                        }
                        else
                            r += string.Format("Процесс Excel {0} был завершен пользователем{1}", pData.procId, Environment.NewLine);
                    }
            }
            else
            {
                foreach (var pData in pids)
                    if (pData.type == 0)
                        if (Process.GetProcesses().Any(p => p.Id == pData.procId))
                            pids_return.Add(pData);
                        else
                            r += string.Format("Процесс Excel {0} был завершен пользователем{1}", pData.procId, Environment.NewLine);

            }

            pids = pids_return;


            return r;

        }
    }
}
