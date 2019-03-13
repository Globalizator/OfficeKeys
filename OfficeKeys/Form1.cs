using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace OfficeKeys
{

    public partial class Form1 : Form
    {
        //Если долго никто не нажимал на клавиши - сохранить документы офиса
        //Последнее нажатие на клавишу в системе
        private static DateTime LastKeyPress;

        //Глобальные хуки на клавиши
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        //Управление окнами
                [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("User32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);
        const int WM_SETTEXT = 0X000C;


        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                //Console.WriteLine((Keys)vkCode);
                LastKeyPress = DateTime.Now;
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }


        public Form1()
        {
            InitializeComponent();
            _hookID = SetHook(_proc);
        }


        private void SaveOffice()
        {
            //Process p = Process.GetProcessesByName("notepad").FirstOrDefault();
            //MessageBox.Show(p.MainWindowHandle.ToString());

            string[] officeapparray = new string[] { "winword", "excel", "project", "powerpoint", "visio", "notepad" };
            foreach (string officeapp in officeapparray)
            {
                foreach (Process q in Process.GetProcessesByName(officeapp))
                {
                    SendKeys.Flush();
                    SetForegroundWindow(q.MainWindowHandle);
                    ShowWindow(q.MainWindowHandle, 1);
                    SetForegroundWindow(q.MainWindowHandle);
                    System.Threading.Thread.Sleep(50);
                    SendKeys.Send("^s");
                }
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            LastKeyPress = DateTime.Now;
        }

        private void tmrMain_Tick(object sender, EventArgs e)
        {
            //получаем разницу во времени между последним нажатием клавиш
            TimeSpan diff = DateTime.Now - LastKeyPress;
            lblSeconds.Text = "Seconds:"+(30-diff.Seconds).ToString();
            //если больше 30 секунд
            if (diff.Seconds >= 30)
            {
                //сброс счётчика секунд
                LastKeyPress = DateTime.Now;
                //сохранение документов
                //MessageBox.Show("Call Save Office");
                SaveOffice();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnhookWindowsHookEx(_hookID);
        }
    }

}



   