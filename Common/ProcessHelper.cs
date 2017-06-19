using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public static class ProcessHelper
    {
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        /// set focus on a process
        /// </summary>
        /// <param processname></param>
        public static void SetFocusOnProcess(string processName)
        {
            Process process = Process.GetProcessesByName(processName).FirstOrDefault();
            SetForegroundWindow(process.MainWindowHandle);
        }

        /// <summary>
        /// kill all the process with processName
        /// </summary>
        /// <param name="processName"></param>
        public static void KillAllProcess(string processName)
        {
            //kill all the processes
            foreach (Process process in Process.GetProcessesByName(processName))
            {
                process.Kill();
            }

        }
    }
}
