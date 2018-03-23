using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace videoToArcive
{
    class Program
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
        [DllImport("User32.dll")]
        static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);
        [DllImport("user32.dll")]
        public static extern int SendMessageA(IntPtr hWnd, uint Msg, int wParam, IntPtr lParam);

        const int WM_SETTEXT = 0x000c;
        const uint WM_KEYDOWN = 0x0100;
        const int LVM_GETITEMCOUNT = 0x1004;
        const uint LVIF_TEXT = 0x00000001;
        const int LVM_GETITEM = 0x1005;

        [StructLayoutAttribute(LayoutKind.Sequential)]
        private struct LVITEM
        {
            public uint mask;
            public int iItem;
            public int iSubItem;
            public uint state;
            public uint stateMask;
            public IntPtr pszText;
            public int cchTextMax;
            public int iImage;
            public IntPtr lParam;
        }
        static void Main(string[] args)
        {
            string sevenZ = @"c:\7ZIP\7zFM.exe";
            if (File.Exists(sevenZ))
            {
                Process.Start(sevenZ);//Запускаем архиватор
                Thread.Sleep(1000);
                IntPtr h1 = (IntPtr)0;
                string caption = "";
                //List<int> dir = new List<int> { 0x2E, 0x43, };
                foreach (System.Diagnostics.Process anti in System.Diagnostics.Process.GetProcesses()) // перебираем все процесы
                {
                    if (anti.ProcessName.Contains("7zFM")) // находим окно по точному заголовку окна
                    {
                        h1 = anti.MainWindowHandle;
                        caption = anti.MainWindowTitle;
                    }
                }
                if (h1!=null)
                {
                    IntPtr h2 = FindWindowEx(h1,new IntPtr(0), "7-Zip::Panel", "");
                    IntPtr h4 = FindWindowEx(h2, new IntPtr(0), "SysListView32", "");
                    IntPtr h3 = FindWindowEx(h2, new IntPtr(0), "ReBarWindow32", "");
                    h2 = FindWindowEx(h3, new IntPtr(0), "ComboBoxEx32", "");
                    h3 = FindWindowEx(h2, new IntPtr(0), "ComboBox", "");
                    h2 = FindWindowEx(h3, new IntPtr(0), "Edit", "");
                    PostMessage(h2, WM_KEYDOWN, 0x2E, 0);
                    Thread.Sleep(500);
                    SendMessage(h2, WM_SETTEXT, 0, @"C:\temp\");
                    Thread.Sleep(500);
                    PostMessage(h2, WM_KEYDOWN, 13, 0);
                    Thread.Sleep(1000);
                    int count = SendMessage(h4, LVM_GETITEMCOUNT, 0, "0");
                    LVITEM lvi = new LVITEM();
                    lvi.mask = LVIF_TEXT;
                    lvi.cchTextMax = 255;
                    lvi.iItem = 50;            // the zero-based index of the ListView item 
                    lvi.iSubItem = 0;
                    lvi.pszText = Marshal.AllocHGlobal(255);
                    IntPtr ptrLvi = Marshal.AllocHGlobal(Marshal.SizeOf(lvi));
                    Marshal.StructureToPtr(lvi, ptrLvi, false);
                    SendMessageA(h4, LVM_GETITEM, 2, ptrLvi);  // Extract the text of the specified item 
                    string itemText = Marshal.PtrToStringAuto(lvi.pszText);
                }
                else
                {
                    Console.WriteLine(sevenZ + " не запустился");
                }
            }
            else
            {
                Console.WriteLine(sevenZ + " не найден");
            }
            Console.WriteLine("Для завершения нажмите любую клавишу");
            Console.ReadKey();
        }
    }
}
