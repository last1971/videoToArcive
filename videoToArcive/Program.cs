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
        const int WM_SETTEXT = 0x000c;
        const uint WM_KEYDOWN = 0x0100;
        
        static void Main(string[] args)
        {
            string sevenZ = @"c:\7ZIP\7zFM.exe";
            if (File.Exists(sevenZ))
            {
                Process p = Process.Start(sevenZ);//Запускаем архиватор
                Thread.Sleep(1000);
                IntPtr h1 = p.MainWindowHandle;
                string caption = p.MainWindowTitle;
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
                    Thread.Sleep(2000);
                    
                    AutomationElement el = AutomationElement.FromHandle(h4);

                    // Walk the automation element tree using content view, so we only see
                    // list items, not scrollbars and headers. (Use ControlViewWalker if you
                    // want to traverse those also.)
                    TreeWalker walker = TreeWalker.ContentViewWalker;
                    int i = 0;
                    Boolean flag = false;
                    for (AutomationElement child = walker.GetFirstChild(el);
                        child != null;
                        child = walker.GetNextSibling(child))
                    {
                        // Print out the type of the item and its name
                        if (child.Current.Name == "video.mp4")
                        {
                            SelectListItem(child);
                            flag = true;
                            break;
                        }
                    }
                    if (flag)
                    {
                        h2 = FindWindowEx(h1, new IntPtr(0), "ToolbarWindow32", "");
                        el = AutomationElement.FromHandle(h2);
                        AutomationElement child = walker.GetFirstChild(el);
                        InvokePattern ptrnServiceRequestTab = child.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        ptrnServiceRequestTab.Invoke();
                        Thread.Sleep(1000);
                        h2 = FindWindow("#32770", "Добавить к архиву");
                        h3 = FindWindowEx(h2, new IntPtr(0), "ComboBox", "");
                        h4 = FindWindowEx(h3, new IntPtr(0), "Edit", "");
                        SendMessage(h4, WM_SETTEXT, 0, "архив.zip");
                        h4 = FindWindowEx(h2, h3, "ComboBox", "");
                        PostMessage(h4, WM_KEYDOWN, 90, 0);
                        Thread.Sleep(1000);
                        h3 = FindWindowEx(h2, new IntPtr(0), "Button", "OK");
                        PostMessage(h3, WM_KEYDOWN, 13, 0);
                        Console.WriteLine("Архив создан");
                        Thread.Sleep(2000);
                        p.Kill();
                    }
                    else
                    {
                        Console.WriteLine("Not find video.mp4");
                    }
                    
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
