using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace videoToArcive
{
    class Program
    {

        static public void SelectListItem( AutomationElement selectionContainer)
        {
            if ((selectionContainer == null))
            {
                throw new ArgumentException(
                    "Argument cannot be null or empty.");
            }

            else
            {
                try
                {
                    SelectionItemPattern selectionItemPattern;
                    selectionItemPattern =
                    selectionContainer.GetCurrentPattern(
                    SelectionItemPattern.Pattern) as SelectionItemPattern;
                    selectionItemPattern.Select();
                }
                catch (InvalidOperationException)
                {
                    // Unable to select
                    return;
                }
            }
        }


        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
        [DllImport("User32.dll")]
        static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);
        [DllImport("user32.dll")]
        static extern int SendMessageA(IntPtr hWnd, uint Msg, int wParam, IntPtr lParam);
        [DllImport( "user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, int dwProcessId);
        [DllImport("kernel32.dll")]
        static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress,uint dwSize, uint flAllocationType, uint flProtect);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, int dwSize, ref int lpNumberOfBytesWritten);
        [DllImport("kernel32.dll")]
        public static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, int dwSize, ref int lpNumberOfBytesRead);
        const int WM_SETTEXT = 0x000c;
        const uint WM_KEYDOWN = 0x0100;
        const int LVM_GETITEMCOUNT = 0x1004;
        const uint LVIF_TEXT = 0x00000001;
        const int LVM_GETITEM = 0x1005;
        const int PROCESS_VM_WRITE = 0x0020;
        const int PROCESS_VM_OPERATION = 0x0008;
        const int MEM_COMMIT = 0x00001000;
        const int PAGE_READWRITE = 0x04;
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


                    AutomationElement el = AutomationElement.FromHandle(h4);

                    // Walk the automation element tree using content view, so we only see
                    // list items, not scrollbars and headers. (Use ControlViewWalker if you
                    // want to traverse those also.)
                    TreeWalker walker = TreeWalker.ContentViewWalker;
                    int i = 0;
                    for (AutomationElement child = walker.GetFirstChild(el);
                        child != null;
                        child = walker.GetNextSibling(child))
                    {
                        // Print out the type of the item and its name
                        if (child.Current.Name == "video.mp4")
                        {
                            SelectListItem(child);
                            Console.WriteLine("item {0} is a \"{1}\" with name \"{2}\"", i++, child.Current.LocalizedControlType, child.Current.Name);
                            break;
                        }
                    }









                    int count = SendMessage(h4, LVM_GETITEMCOUNT, 0, "0");
                    int hz = 0;
                    GetWindowThreadProcessId(h4, out hz);
                    IntPtr hProcess = OpenProcess(PROCESS_VM_WRITE | PROCESS_VM_OPERATION, true, hz);
                    
                    LVITEM lvi = new LVITEM();
                    lvi.mask = LVIF_TEXT;
                    lvi.cchTextMax = 512;
                    lvi.iItem = 0;            // the zero-based index of the ListView item 
                    lvi.iSubItem = 0;
                    //lvi.pszText = Marshal.AllocHGlobal(255);
                   
                    int cdWrite = 0;
                    byte[] managedArray = new byte[Marshal.SizeOf(lvi)];
                    IntPtr ptrLvi = Marshal.AllocHGlobal(Marshal.SizeOf(lvi));
                    //Marshal.StructureToPtr(lvi, ptrLvi, false);
                    Marshal.Copy(ptrLvi, managedArray, 0, Marshal.SizeOf(lvi));
                    IntPtr lpMem = VirtualAllocEx(hProcess, new IntPtr(0), (uint)Marshal.SizeOf(lvi), MEM_COMMIT, PAGE_READWRITE);
                    IntPtr text = VirtualAllocEx(hProcess, new IntPtr(0), 512, MEM_COMMIT, PAGE_READWRITE);
                    lvi.pszText = text;
                    WriteProcessMemory(hProcess, lpMem, managedArray, Marshal.SizeOf(lvi), ref cdWrite);
                    SendMessageA(h4, LVM_GETITEM, 0, lpMem);  // Extract the text of the specified item 
                    int bytesRead = 0;
                    byte[] buffer = new byte[512]; //'Hello World!' takes 12*2 bytes because of Unicode 
                    ReadProcessMemory(hProcess, text, buffer, buffer.Length, ref bytesRead);
                    string a = Marshal.PtrToStringUni(text, 512);
                    //string itemText = Marshal.PtrToStringAuto(lvi.pszText);
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
