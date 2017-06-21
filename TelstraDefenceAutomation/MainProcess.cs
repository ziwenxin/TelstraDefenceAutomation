using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Remoting.Contexts;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using WindowsInput;
using WindowsInput.Native;
using BusinessObjects;
using BusinessObjects.MERIDIAN;
using BusinessObjects.SharePoint;
using Common;
using Exceptions;
using ICSharpCode.SharpZipLib.Tar;
using Microsoft.Win32.TaskScheduler;
using NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;
using WinSCP;
using Exception = System.Exception;


namespace TelstraDefenceAutomation
{
    public class MainProcess
    {

        //log string
        private static StringBuilder sb=new StringBuilder();
        //static member to store config file
        private static ISheet configSheet;
        static void Main(string[] args)
        {
            #region MainProcess
            int retryTimes = 0;
            try
            {
                //kill all the excel process
                ProcessHelper.KillAllProcess("EXCEL");
                //read settings and set default download folder for chrome
                configSheet = Intialization();
                //get retry times
                retryTimes = (int)configSheet.GetRow(21).GetCell(1).NumericCellValue;

                //before automation, delete all files in the save folder
                DeleteAllFiles(configSheet.GetRow(5).GetCell(1).StringCellValue);

                //download excel files
                DownLoadTollDocuments();
                DownLoadMeridianDocuments();

                //delete several lines at the beginning
                ProcessExcels();

                //copy files from share folder
                DownLoadShareFolderDocs();

                //download files from share point
                DownLoadSharePointDoc();
                //upload to server
                UploadFiles();

                //run lavastorm program
                RunLavaStorm();
                //renew retry times
                retryTimes = 3;

            }
            catch (Exception e)
            {
                //if still needs to retry
                if (retryTimes > 0)
                {
                    //reschedule one run
                    RescheduleTask();
                    retryTimes--;
                }
                //notify admin
                else
                {
                    //reset retry times
                    retryTimes = 3;
                    SendEmail();
                }
                //log
                AddToLog(e.ToString());
            }
            finally
            {
                try
                {
                    //reset retry times
                    configSheet.GetRow(21).GetCell(1).SetCellValue(retryTimes);
                    //delete the previous file

                    if (File.Exists("Defense Automation Config.xlsx"))
                        File.Delete("Defense Automation Config.xlsx");
                    //write back to config file 
                    ExcelHelper.Save(configSheet, "Defense Automation Config.xlsx");
                    //write log file
                    WriteLogFile();
                }
                catch (Exception e)
                {
                    AddToLog(e.Message);
                }
            }





            Exit();
            #endregion





        }

        /// <summary>
        /// run lavastrom to process the excel files in the server side
        /// </summary>
        private static void RunLavaStorm()
        {
            //new a process to open the file
            using (Process proc = new Process())
            {
                //log 
                AddToLog("Start to run lavastorm...");
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
                string folder = configSheet.GetRow(36).GetCell(1).StringCellValue;
                string filename = configSheet.GetRow(37).GetCell(1).StringCellValue;

                string dosLine = @"D:\Users\D795314\DATA_IMPORT38_for_test.brg";
                proc.StandardInput.WriteLine(dosLine);
                //exit
                proc.StandardInput.WriteLine("exit");
                //wait for the application appears

                Thread.Sleep(60000);
                //set focus on the window
                ProcessHelper.SetFocusOnProcess("bre");
                Thread.Sleep(5000);
                //input simulator
                InputSimulator simulator = new InputSimulator();
                //move the mouse
                simulator.Mouse.MoveMouseTo(33000, 30000);
                simulator.Mouse.LeftButtonClick();
                Thread.Sleep(1000);

                //select all the process
                simulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_A);
                //click rerun
                Thread.Sleep(1000);
                //move the mouse
                simulator.Mouse.MoveMouseTo(24500, 22500);
                simulator.Mouse.LeftButtonClick();

                //wait for running
                Thread.Sleep(60000);
                //save the programs
                simulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_S);
                Thread.Sleep(3000);
                //kill the process
                ProcessHelper.KillAllProcess("bre");
                //log
                AddToLog("Lavastrom runs completed");
            }




        }


        

        /// <summary>
        /// send a email to the address in config file, using the current outlook account
        /// </summary>
        private static void SendEmail()
        {
            //set address, subject and body of email
            string emailAddr = configSheet.GetRow(22).GetCell(1).StringCellValue;
            string subject = "Automation Rerun Failed";
            string autoPath = configSheet.GetRow(19).GetCell(1).StringCellValue;
            string body = "Please go to desktop to run the TelstraDefenceAutomation.exe manually."+Environment.NewLine+" Alternatively, you can go to '" + autoPath + "' to run TelstraDefenceAutomation.exe manually."+ Environment.NewLine + " Thanks";
            //set up a mail
            Application app = new Application();
            MailItem mail = (MailItem)app.CreateItem(OlItemType.olMailItem);

            mail.To = emailAddr;
            mail.Body = body;
            mail.Subject = subject;
            //set up account
            Accounts accs = app.Session.Accounts;
            Account acc = null;
            foreach (Account account in accs)
            {
                acc = account;
                break;
            }
            mail.SendUsingAccount = acc;
            //send email
            mail.Send();
        }

        /// <summary>
        /// download 'Deployment Planning and Tracking' from share point
        /// </summary>
        private static void DownLoadSharePointDoc()
        {
            AddToLog("Downloading from share point...");
            //download file from share point
            SharePointPage sharePointPage = new SharePointPage(configSheet);
            sharePointPage.DownLoadSharePointDoc();
            //change 1 sheet name from BV & SA to BVSA
            //get path and filename
            string savepath = configSheet.GetRow(5).GetCell(1).StringCellValue;
            string filename = configSheet.GetRow(34).GetCell(1).StringCellValue;
            savepath += "\\";
            //set sheet name
            OfficeExcel.ChangeSheetName(savepath, filename, "BV & SA", "BVSA");

            AddToLog("DownLoad from share point completed");
        }

        /// <summary>
        /// download 'Logistics','All-CECs-StockTransfer Burwood' and 'All-CECs-StockTransfer-Regents' from share folder
        /// </summary>
        private static void DownLoadShareFolderDocs()
        {
            AddToLog("Downloading files from share folder...");

            //get username and password
            string username = configSheet.GetRow(24).GetCell(1).StringCellValue;
            string password = configSheet.GetRow(25).GetCell(1).StringCellValue;
            //get local save path and server path
            string localPath = configSheet.GetRow(5).GetCell(1).StringCellValue;
            string serverPath = configSheet.GetRow(26).GetCell(1).StringCellValue;
            //get filename
            string filename = configSheet.GetRow(27).GetCell(1).StringCellValue;
            localPath += "\\";
            //launch a command line to connect to the server
            ConnectState(serverPath, username, password);

            //copy logistic schedule file
            filename += ".xlsx";
            serverPath += "\\";
            File.Copy(serverPath + filename, localPath + filename, true);
            AddToLog(filename + " download completed");
            //copy Burwood stock transfer file
            serverPath = configSheet.GetRow(28).GetCell(1).StringCellValue + "\\";
            filename = configSheet.GetRow(29).GetCell(1).StringCellValue;
            filename = GetNewestFileName(serverPath, filename);
            File.Copy(serverPath + filename, localPath + filename, true);
            AddToLog(filename + " download completed");

            //copy Regents transfer stock file
            //copy Burwood stock transfer file
            serverPath = configSheet.GetRow(30).GetCell(1).StringCellValue + "\\";
            filename = configSheet.GetRow(31).GetCell(1).StringCellValue;
            filename = GetNewestFileName(serverPath, filename);
            File.Copy(serverPath + filename, localPath + filename, true);
            AddToLog(filename + " download completed");
        }

        /// <summary>
        /// upload all the files to the server using WinScp
        /// </summary>
        private static void UploadFiles()
        {
            AddToLog("Start to upload files to server...");
            //setup session options
            SessionOptions options = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = configSheet.GetRow(13).GetCell(1).StringCellValue,
                UserName = configSheet.GetRow(14).GetCell(1).StringCellValue,
                Password = configSheet.GetRow(15).GetCell(1).StringCellValue,
                SshHostKeyFingerprint = configSheet.GetRow(16).GetCell(1).StringCellValue,
            };

            using (Session session = new Session())
            {
                //connect
                session.Open(options);

                //upload files
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;

                //get path
                string localPath = configSheet.GetRow(5).GetCell(1).StringCellValue;
                string remotePath = configSheet.GetRow(17).GetCell(1).StringCellValue;
                localPath += "\\";
                //change the '/' to '\'
                localPath = localPath.Replace("/", "\\");

                //upload the files into server,delete the files in the local
                TransferOperationResult operationResult =
                    session.PutFiles(localPath + "*.xls*", remotePath, true, transferOptions);

                //throw any error
                operationResult.Check();

                //print result
                foreach (TransferEventArgs transfer in operationResult.Transfers)
                {
                    AddToLog(string.Format("Upload of {0} successed", transfer.FileName));
                }
                AddToLog("Upload all completed");
            }

        }

        /// <summary>
        /// process the excel files downloaded from 'Toll' and 'Meridian', it mainly delete several lines from the top of the documents
        /// </summary>
        private static void ProcessExcels()
        {
            AddToLog("Processing Excel files...");

            //get total report numbers
            int totalWaitMilliSecs = 0;
            int totalReportNum = (int)configSheet.GetRow(6).GetCell(1).NumericCellValue;
            for (int i = 0; i < totalReportNum; i++)
            {
                //read from report
                string savePath = configSheet.GetRow(5).GetCell(1).StringCellValue;
                string filename = configSheet.GetRow(7).GetCell(1 + i).StringCellValue;
                savePath += "\\";                
                string filepath = savePath + filename;
                //check if the file exists
                string extension = ".xlsx";

                if (!File.Exists(filepath + extension))
                {
                    if (!File.Exists(filepath + ".xls"))
                        throw new Exception(filepath + " is not downloaded");
                }
                int linesToBeDeleted = (int)configSheet.GetRow(8).GetCell(1 + i).NumericCellValue;

                //use library to read an excel file
                try
                {
                    ISheet reportsheet = ExcelHelper.ReadExcel(filepath + extension);

                    //do the archive
                    ExcelHelper.MoveFileToArchive(savePath, filename, ".xlsx");
                    //save
                    ExcelHelper.SaveTo(reportsheet, filepath + ".xlsx", linesToBeDeleted);
                }
                catch (Exception e)
                {

                    //get the rename
                    string rename = configSheet.GetRow(6).GetCell(1 + i).StringCellValue;
                    //process the file by string
                    ExcelHelper.ProcessInvalidExcel(savePath, filename, rename);
                    //save the incorrupted file as xlsx
                    OfficeExcel.SaveAs(savePath, rename);
                    //delete he priginal file
                    if (File.Exists(savePath + rename + ".xls"))
                        File.Delete(savePath + rename + ".xls");
                }
                
                AddToLog(filename + " process completed");
            }



        }

        /// <summary>
        /// delete all the files in the local save path, excluding the folder
        /// </summary>
        /// <param name="path"></param>
        private static void DeleteAllFiles(string path)
        {
            //log
            AddToLog("Deleting all previous files...");
            //get directory info
            DirectoryInfo di = new DirectoryInfo(path);
            foreach (FileInfo fileInfo in di.GetFiles())
            {
                fileInfo.Delete();
            }
            //log
            AddToLog("Delete completed");
        }

        /// <summary>
        /// initial the webdriver and read data from config file
        /// </summary>
        /// <returns>the work sheet of config file</returns>
        private static ISheet Intialization()
        {
            //log
            AddToLog("Inialising...");

            int retryCount = 3;
            //read data
            ISheet sheet = ExcelHelper.ReadExcel("Defense Automation Config.xlsx");

            //check if the download folder exists, if not create one
            string path = sheet.GetRow(5).GetCell(1).StringCellValue;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            //change default download location
            while (true)
            {
                try
                {
                    WebDriver.Init(path);
                    break;
                }
                //retry at most 3 times to initalize the driver
                catch (Exception e)
                {
                    //close the previous window
                    if (WebDriver.ChromeDriver != null)
                        WebDriver.ChromeDriver.Quit();
                    if (retryCount <= 0)
                    {
                        throw e;
                    }
                    retryCount--;
                }
            }
            //log
            AddToLog("Initialization completed");
            return sheet;
        }

        /// <summary>
        /// download 'TelDef - Shipped Order Report','TelDef - SOH Detail' and 'TelDef - Goods Receipt By Date Range' from toll 
        /// </summary>
        private static void DownLoadTollDocuments()
        {
            AddToLog("DownLoading documents from Toll...");

            //login
            try
            {
                TollLoginPage tollLoginPage = new TollLoginPage(configSheet);
                TollReportDownloadPage tollDownloadPage = tollLoginPage.Login();
                string savepath = configSheet.GetRow(5).GetCell(1).StringCellValue;
                savepath += "\\";
                //download first document
                string filename = configSheet.GetRow(7).GetCell(1).StringCellValue;
                TollGoodReportPage tollGoodReportPage = tollDownloadPage.DownloadGoodDocument();
                tollGoodReportPage.DownLoadReport(savepath + filename);

                AddToLog(filename + " download completed");
                //download 2nd
                filename = configSheet.GetRow(7).GetCell(2).StringCellValue;
                tollDownloadPage.GoToReportPage();
                TollShipOrderPage tollShipDetailPage = tollDownloadPage.DownLoadShipOrder();
                tollShipDetailPage.DownLoadReport(savepath + filename);

                AddToLog(filename + " download completed");

                //download the 3rd 
                filename = configSheet.GetRow(7).GetCell(3).StringCellValue;
                tollDownloadPage.GoToReportPage();
                TollSOHDetailPage tollSohDetailPage = tollDownloadPage.DownloadSOHDetail();
                tollSohDetailPage.DownLoadReport(savepath + filename);

                AddToLog(filename + " download completed");

            }
            catch (NoReportsException reportsException)
            {
                throw new Exception("The 'TollTotalDocuments' cell could not be empty");
            }
            catch (Exception e)
            {
                throw e;

            }

        }

        /// <summary>
        /// download 'PO_DETAILS_REPORT' and 'Accounting_Details_from_meridian' from Meridian
        /// </summary>
        private static void DownLoadMeridianDocuments()
        {
            AddToLog("Downloading Meridian documents...");
            //go to the portal of meridian
            MeridianPortalPage meridianPortalPage = new MeridianPortalPage(configSheet);
            MeridianNavigationPage meridianNavigationPage = meridianPortalPage.LaunchMeridian();

            //download files
            //if it fails, retry it
            int retryCount = 3;
            while (true)
            {
                try
                {
                    DownLoadPODetailDoc(configSheet, meridianNavigationPage);
                    break;
                }
                catch (Exception e)
                {
                    //if file exists, delete it
                    string savepath = configSheet.GetRow(5).GetCell(1).StringCellValue;
                    string filename = configSheet.GetRow(7).GetCell(4).StringCellValue;
                    savepath += "\\";
                    DeleteFile(savepath, filename);
                    if (retryCount <= 0)
                        throw e;
                    retryCount--;
                    //switch back to main frame
                    WebDriver.ChromeDriver.SwitchTo().DefaultContent();
                }
            }
            //retry
            retryCount = 3;
            while (true)
            {
                try
                {
                    DownLoadAccDetailDoc(configSheet, meridianNavigationPage);
                    break;
                }
                catch (Exception e)
                {
                    //if file exists, delete it
                    string savepath = configSheet.GetRow(5).GetCell(1).StringCellValue;
                    string filename = configSheet.GetRow(7).GetCell(5).StringCellValue;
                    DeleteFile(savepath, filename);
                    if (retryCount <= 0)
                        throw e;
                    retryCount--;
                    //switch back to main frame
                    WebDriver.ChromeDriver.SwitchTo().DefaultContent();
                }
            }
            //log
            AddToLog("Download Meridian Documents completed.");
        }

        /// <summary>
        /// download 'Accounting_Details_from_meridian'
        /// </summary>
        /// <param name="configSheet"></param>
        /// <param name="meridianNavigationPage"></param>
        private static void DownLoadAccDetailDoc(ISheet configSheet, MeridianNavigationPage meridianNavigationPage)
        {
            //go to account payable entry detail page
            MeridianVariableEntryPage accVariableEntryPage = meridianNavigationPage.GotoAccountDetailEntryPage(configSheet);
            MeridianAccountDetailPage accountDetailPage = accVariableEntryPage.AccountEnter();
            ////click open button select detail
            MeridianPopUpWindow accPopUpWindow = accountDetailPage.OpenPoPUpWindow();
            accPopUpWindow.SelectAccountDetailDoc();
            //get full path
            string filename = configSheet.GetRow(7).GetCell(5).StringCellValue;
            string savepath = configSheet.GetRow(5).GetCell(1).StringCellValue;
            savepath += "\\";
            
            ////download PO Detail Reprrt
            accountDetailPage.DownLoadAccountDetailDoc(savepath + filename);
           
            AddToLog(filename + " download completed");
        }

        /// <summary>
        /// download 'PO_DETAILS_REPORT'
        /// </summary>
        /// <param name="configSheet"></param>
        /// <param name="meridianNavigationPage"></param>
        private static void DownLoadPODetailDoc(ISheet configSheet, MeridianNavigationPage meridianNavigationPage)
        {
            //go to PO detail
            MeridianVariableEntryPage POVariableEntryPage = meridianNavigationPage.GotoPoDetailEntryPage(configSheet);
            MeridianAccountDetailPage PoAccountDetailPage = POVariableEntryPage.PODetailEnter();
            //click open button select detail
            MeridianPopUpWindow POPopUpWindow = PoAccountDetailPage.OpenPoPUpWindow();
            POPopUpWindow.SelectPODetailDoc();
            //get full path
            string filename = configSheet.GetRow(7).GetCell(4).StringCellValue;
            string savepath = configSheet.GetRow(5).GetCell(1).StringCellValue;
            savepath += "\\";
            //download PO Detail Reprrt
            PoAccountDetailPage.DownLoadPoDetailDoc(savepath + filename);

            AddToLog(filename + " download completed");

        }

        /// <summary>
        /// reschedule the program after 1 min
        /// </summary>
        public static void RescheduleTask()
        {
            //get the service on the local 
            using (TaskService ts = new TaskService())
            {
                string taskName = "Rerun";

                //create a new task
                TaskDefinition td = ts.NewTask();
                td.RegistrationInfo.Description = "Redo the automation";
                //set expires date
                td.Settings.DeleteExpiredTaskAfter = TimeSpan.FromSeconds(10);
                //create a trigger that execute this task 1 minutes later
                td.Triggers.Add(new TimeTrigger { EndBoundary = DateTime.Now.AddMinutes(2) });
                td.Triggers.Add(new TimeTrigger(DateTime.Now.AddMinutes(1)));
                //get automation path
                string autoPath = configSheet.GetRow(19).GetCell(1).StringCellValue;


                //create an action
                td.Actions.Add(new ExecAction(autoPath + "TelstraDefenceAutomation.exe", null, autoPath));
                //register the task in the root folder
                ts.RootFolder.RegisterTaskDefinition(taskName, td);
            }


        }

        /// <summary>
        /// exit the program
        /// </summary>
        private static void Exit()
        {
            AddToLog("The automation will be closed in 5 secs");
            //close the automation
            Thread.Sleep(5000);
            if (WebDriver.ChromeDriver != null)
                WebDriver.ChromeDriver.Quit();
            Environment.Exit(0);
        }

        /// <summary>
        /// connect to the remote share folder
        /// </summary>
        /// <param name="path">share folder path</param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns> if connect success</returns>
        private static bool ConnectState(string path, string username, string password)
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

        /// <summary>
        /// find the newest version of file in the share folder
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="filename"></param>
        /// <returns>the newest file name</returns>
        private static string GetNewestFileName(string filepath, string filename)
        {

            //get directory info
            DirectoryInfo directory = new DirectoryInfo(filepath);
            //get the latest file
            return directory.GetFiles(filename + "*.xlsx").OrderByDescending(f => f.LastWriteTime).First().Name;
        }
        /// <summary>
        /// Delete a file if exists
        /// </summary>
        /// <param name="savepath"></param>
        /// <param name="filename"></param>
        private static void DeleteFile(string savepath, string filename)
        {
            string fullPath = savepath + filename + ".xls";
            if (File.Exists(fullPath))
                File.Delete(fullPath);
        }

        /// <summary>
        /// print message to output window and log the message
        /// </summary>
        /// <param name="msg"></param>
        private static void AddToLog(string msg)
        {
            sb.Append(msg + "\r\n");
            Console.WriteLine(msg);
        }

        private static void WriteLogFile()
        {
            //get path and file name
            string path = configSheet.GetRow(20).GetCell(1).StringCellValue;
            //create folder
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            //create folder of current date
            path += "\\"+DateTime.Today.ToString("d").Replace("/","-");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            string filename = path + "\\" + DateTime.Now.ToString().Replace("/", "-").Replace(":"," ")+".log";
            //write log file
            File.WriteAllText(filename,sb.ToString());
        }
    }


}

