using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenM
{
    static public class HKEY
    {
        public static bool checkMachineType(int type)
        {
            if (type == 1)
            {
                RegistryKey winLogonKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Word\Addins\WordAddin2", true);
                string currentKey = winLogonKey.GetValue("LoadBehavior").ToString();

                if (currentKey == "0" || currentKey == "2" || currentKey == "8")
                    return (false);
                return (true);
            }
            if (type == 0)
            {
                RegistryKey winLogonKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Excel\Addins\ExcelAddin2\", true);
                string currentKey = winLogonKey.GetValue("LoadBehavior").ToString();

                if (currentKey == "0" || currentKey == "2" || currentKey == "8")
                    return (false);
                return (true);
            }
            else return (false);
        }

        public static void AddReg(int ID)
        {
            
            RegistryKey winLogonKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Settings\", true);
            winLogonKey.SetValue("AppID", ID, RegistryValueKind.DWord);

        }


        public static string GetRegistryValue(int type,string parametr)                             // получаем параметры из регистра на основе типа надстройки, и названия параметра в регистре
        {
            
            if (type == 1)
            {
                RegistryKey winLogonKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Settings\", true);
                string Value = winLogonKey.GetValue(parametr).ToString();

                
                return (Value);
            }
            if (type == 0)
            {
                RegistryKey winLogonKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Settings\", true);
                string Value = winLogonKey.GetValue(parametr).ToString();

                
                return (Value);
            }
            else return ("");
        }
        public static int GetRegistryValue(string parametr)                                                     //для получения айди процесса из регистра
        {
            RegistryKey winLogonKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Settings\", true);
            int? Value = winLogonKey.GetValue(parametr) as int?;
            int Value1;
            Value1 = Value ?? default(int);
            return Value1;
        }
    }
}
