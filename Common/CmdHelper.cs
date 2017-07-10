using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WindowsInput;
using WindowsInput.Native;
using NPOI.SS.UserModel;

namespace Common
{
    public static class CmdHelper
    {
        /// <summary>
        /// run lavastrom to process the excel files in the server side
        /// </summary>
        public static void RunLavaStorm(Dictionary<string,string> configDic)
        {
            //new a process to open the file
            using (Process proc = new Process())
            {

                //set parameters
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                //start and input
                proc.Start();
                //get program folder and name
                string folder = configDic["ProgramFolder"];
                string filename = configDic["ProgramName"];

                string dosLine = "\"" + folder + "\\" + filename + "\"";
                proc.StandardInput.WriteLine(dosLine);
                //exit
                proc.StandardInput.WriteLine("exit");
                //wait for the application appears

                Thread.Sleep(120000);
                //set focus on the window
                ProcessHelper.SetFocusOnProcess("bre");
                Thread.Sleep(5000);
                //input simulator
                InputSimulator simulator = new InputSimulator();
                //move the mouse
                simulator.Mouse.MoveMouseTo(33000, 30000);
                Thread.Sleep(1000);
                simulator.Mouse.LeftButtonClick();
                Thread.Sleep(1000);

                //select all the process
                simulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_A);
                //click rerun
                Thread.Sleep(1000);
                //move the mouse
                simulator.Mouse.MoveMouseTo(24500, 22500);
                Thread.Sleep(1000);
                simulator.Mouse.LeftButtonClick();

                //wait for running
                Thread.Sleep(120000);
                //save the programs
                simulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_S);
                Thread.Sleep(3000);
                //kill the process
                ProcessHelper.KillAllProcess("bre");

            }


    }
        /// <summary>
        /// connect to the remote share folder
        /// </summary>
        /// <param name="path">share folder path</param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns> if connect success</returns>
        public static bool ConnectState(string path, string username, string password)
        {
            //connect result
            bool flag = false;
            using (Process proc = new Process())
            {

                //set parameters
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                //start and input
                proc.Start();
                string dosLine = @"net use " + path + " /User:" + username + " " + password + " /PERSISTENT:YES";
                proc.StandardInput.WriteLine(dosLine);
                //exit
                proc.StandardInput.WriteLine("exit");
                //wait for exit
                while (!proc.HasExited)
                {
                    proc.WaitForExit(1000);
                }
                //get error messages
                string errormsg = proc.StandardError.ReadToEnd();
                proc.StandardError.Close();
                if (string.IsNullOrEmpty(errormsg))
                {
                    flag = true;
                }
                else
                {
                    throw new Exception(errormsg);
                }
            }
            return flag;
        }


    }
}
