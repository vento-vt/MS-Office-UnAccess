using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OpenM
{
    public struct SendingParams
    {
        public int appType, actionType, docID;
    }
    public class MPV
    {
        public static event Action<int, int, int> DocEvent;
        static Socket sck;
        static EndPoint epLocal, epRemote;
        static bool connection = false;
        static string IP, CIP, SPort, CPort;

        public static EndPoint getPoint(int type)
        {
            if (type == 0) return epLocal;
            else return epRemote;
        }

        public static byte[] getBytes(SendingParams str)              //для отправки параметров
        {
            int size = Marshal.SizeOf(str);
            byte[] arr = new byte[size];

            IntPtr ptr = Marshal.AllocHGlobal(size);
            Marshal.StructureToPtr(str, ptr, true);
            Marshal.Copy(ptr, arr, 0, size);
            Marshal.FreeHGlobal(ptr);
            return arr;
        }

        static byte[] getBytes(int str)     //для отправки АЙДИ процесса
        {
            int size = Marshal.SizeOf(str);
            byte[] arr = new byte[size];

            IntPtr ptr = Marshal.AllocHGlobal(size);
            Marshal.StructureToPtr(str, ptr, true);
            Marshal.Copy(ptr, arr, 0, size);
            Marshal.FreeHGlobal(ptr);
            return arr;
        }

             public static int fromBytesInt(byte[] arr)    //для получения АЙДИ процесса
       {
           int str = new int();

           int size = Marshal.SizeOf(str);
           IntPtr ptr = Marshal.AllocHGlobal(size);

           Marshal.Copy(arr, 0, ptr, size);

           str = (int)Marshal.PtrToStructure(ptr, str.GetType());
           Marshal.FreeHGlobal(ptr);

           return str;
       }
                   

        static SendingParams fromBytes(byte[] arr)
        {
            SendingParams str = new SendingParams();

            int size = Marshal.SizeOf(str);
            IntPtr ptr = Marshal.AllocHGlobal(size);

            Marshal.Copy(arr, 0, ptr, size);

            str = (SendingParams)Marshal.PtrToStructure(ptr, str.GetType());
            Marshal.FreeHGlobal(ptr);

            return str;
        }

        public static void Sending(SendingParams param)
        {
            try
            {
                //конвертация текста в байты и его передача
                byte[] msg = getBytes(param);

                sck.Send(msg);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        public static void Sending(int ID)
        {
            try
            {
                //конвертация в байты и передача
                byte[] msg = getBytes(ID);
                sck.Send(msg);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        /*   private void Sending(string Text)
        {
            try
            {
                //конвертация текста в байты и его передача
                System.Text.ASCIIEncoding enc = new System.Text.ASCIIEncoding();
                byte[] msg = new byte[1500];

                msg = Encoding.Default.GetBytes(Text);

                sck.Send(msg);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        } */

        public static Socket Connect(int type,int Ptype)
        {

            sck = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
            sck.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            if (!connection)
            {
                connect(type,Ptype);

                connection = true;
            };
            return sck;
        }

        static public void MessageCallBack(IAsyncResult aresut)
        {
            try
            {
                int size = sck.EndReceiveFrom(aresut, ref epRemote);

                if (size > 0)
                {

                    byte[] receivedData = new byte[1500];

                    receivedData = (byte[])aresut.AsyncState;

                    SendingParams received = new SendingParams();

                    received = fromBytes(receivedData);

                    //   ASCIIEncoding eEncoding = new ASCIIEncoding();

                    //   string receivedMessage = Encoding.Default.GetString(receivedData);
                    //  MessageBox.Show(received.actionType.ToString() + received.appType.ToString() + received.docID.ToString());
                    DocEvent?.Invoke(received.actionType, received.appType, received.docID);

                }

                byte[] buffer = new byte[1500];
                sck.BeginReceiveFrom(buffer, 0, buffer.Length, SocketFlags.None, ref epRemote, new AsyncCallback(MessageCallBack), buffer);

            }
            catch (Exception exp)
            {

                MessageBox.Show(exp.ToString());


            }
        }
        

  
        


        public static void GetVal(int Ptype)                                    //присваиваются параметры айпи и портов. Функция из HKEY
        {
             IP = HKEY.GetRegistryValue(Ptype, "IP_Server");
             CIP = HKEY.GetRegistryValue(Ptype, "IP_Client");
             SPort = HKEY.GetRegistryValue(Ptype, "Port_Server");
             CPort = HKEY.GetRegistryValue(Ptype, "Port_Client");
       }


        static private void connect(int SCtype,int Ptype)
        { if (SCtype == 0) { 
            try
            {   if (Ptype == 0)

                    {
                        GetVal(Ptype);              //присваиваются параметры айпи и портов. Функция из HKEY


                        epLocal = new IPEndPoint(IPAddress.Parse(IP), Convert.ToInt32(SPort));
                        sck.Bind(epLocal);

                        epRemote = new IPEndPoint(IPAddress.Parse(CIP), Convert.ToInt32(CPort));
                        sck.Connect(epRemote);

                        byte[] buffer = new byte[1500];
                        sck.BeginReceiveFrom(buffer, 0, buffer.Length, SocketFlags.None, ref epRemote, new AsyncCallback(MPV.MessageCallBack), buffer);
                    }
            else
                    {

                        GetVal(Ptype);              //присваиваются параметры айпи и портов. Функция из HKEY


                        epLocal = new IPEndPoint(IPAddress.Parse(IP), Convert.ToInt32(SPort));
                        sck.Bind(epLocal);

                        epRemote = new IPEndPoint(IPAddress.Parse(CIP), Convert.ToInt32(CPort));
                        sck.Connect(epRemote);

                        byte[] buffer = new byte[1500];
                        sck.BeginReceiveFrom(buffer, 0, buffer.Length, SocketFlags.None, ref epRemote, new AsyncCallback(MPV.MessageCallBack), buffer);
                    }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        else
            {
                try
                {
                    if (Ptype == 1)
                    {
                        GetVal(Ptype);              //присваиваются параметры айпи и портов. Функция из HKEY


                        epLocal = new IPEndPoint(IPAddress.Parse(CIP), Convert.ToInt32(CPort));
                        sck.Bind(epLocal);

                        epRemote = new IPEndPoint(IPAddress.Parse(IP), Convert.ToInt32(SPort));
                        sck.Connect(epRemote);

                       // byte[] buffer = new byte[1500];
                      //  sck.BeginReceiveFrom(buffer, 0, buffer.Length, SocketFlags.None, ref epRemote, new AsyncCallback(MPV.MessageCallBack), buffer);
                    }
                    else
                    {
                        GetVal(Ptype);              //присваиваются параметры айпи и портов. Функция из HKEY

                        epLocal = new IPEndPoint(IPAddress.Parse(CIP), Convert.ToInt32(CPort));
                        sck.Bind(epLocal);
                        
                        epRemote = new IPEndPoint(IPAddress.Parse(IP), Convert.ToInt32(SPort));
                        sck.Connect(epRemote);

                      //  byte[] buffer = new byte[1500];
                     //   sck.BeginReceiveFrom(buffer, 0, buffer.Length, SocketFlags.None, ref epRemote, new AsyncCallback(MPV.MessageCallBack), buffer);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
       }
    }
}
