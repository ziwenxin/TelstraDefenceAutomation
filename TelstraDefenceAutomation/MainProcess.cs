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
using BusinessObjects.Toll;
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
        //dictionary to store all the config data
        private static Dictionary<string, string> _configDic;
        //log string
        private static StringBuilder sb = new StringBuilder();
        //static member to store config file
        static void Main(string[] args)
        {
            #region MainProcess
            int retryTimes = 0;
            ISheet configSheet = null;
            try
            {
                //kill all the excel process
                ProcessHelper.KillAllProcess("EXCEL");
                //read settings and set default download folder for chrome
                configSheet = Intialization();


                //get retry times
                retryTimes = int.Parse(_configDic["RerunTimes"]);

                //before automation, delete all files in the save folder
                DeleteAllFiles(_configDic["LocalSavePath"]);

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
                if (_configDic["EnableAutomation?"].ToLower() == "yes")
                {
                    AddToLog("Start to run lavastorm...");
                    CmdHelper.RunLavaStorm(_configDic);
                    AddToLog("Lavastrom runs completed");
                }

                //renew retry times
                retryTimes = 3;

                //send email
                string subject = "Automation Success";
                SendEmail(subject,"The automation runs successfully on "+DateTime.Now);
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
                    //set content and subject
                    string autoPath = _configDic["AutomationPath"];
                    string subject = "Automation Rerun Failed";
                    string content = "Hello," + Environment.NewLine + Environment.NewLine +
                                     "The bot supporting the Defence Inventory Data Hub has failed to run automatically overnight." +
                                     Environment.NewLine + Environment.NewLine +
                                     "Please go to the desktop to run the executable file manually, as the network or a source may not have been available during the automated run." +
                                     Environment.NewLine + "The executable file is:  TelstraDefenceAutomation.exe" +
                                     Environment.NewLine + Environment.NewLine + "(Alternatively, you can go to '" +
                                     autoPath +
                                     "' to run TelstraDefenceAutomation.exe manually from its saved location)" +
                                     Environment.NewLine +
                                     "If the manual run fails, please refer to the handbook for troubleshooting steps by going to Desktop." +
                                     Environment.NewLine + Environment.NewLine +
                                     "The user handbook file is: Telstra Defence Automation User Handbook v1.*.*.docx" +
                                     Environment.NewLine +
                                     Environment.NewLine + "Thank you for helping me complete my run," +
                                     Environment.NewLine +
                                     "The Defence Inventory Data Hub bot";
                    SendEmail(subject, content);
                }


                //log
                AddToLog(e.ToString());
            }
            finally
            {
                try
                {
                    //reset retry times
                    configSheet.GetRow(24).GetCell(1).SetCellValue(retryTimes);
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
        /// it will read data from config sheet to a dictionary
        /// </summary>
        private static void StoreIntoDic(ISheet configSheet)
        {
            //initial
            _configDic = new Dictionary<string, string>();
            //read line by line
            for (int i = 0; i <= configSheet.LastRowNum; i++)
            {

                //read data in a row
                int value_idx = 1;
                while (configSheet.GetRow(i).GetCell(value_idx) != null && configSheet.GetRow(i).GetCell(value_idx).CellType != CellType.Blank)
                {
                    //get name
                    string name = configSheet.GetRow(i).GetCell(0).StringCellValue;
                    //if the row contains multiple data, name the key with idx
                    if (configSheet.GetRow(i).GetCell(2) != null && configSheet.GetRow(i).GetCell(2).CellType != CellType.Blank)
                        name += value_idx;
                    //if its not string, then convert it to string
                    string value = "";
                    try
                    {
                        value = configSheet.GetRow(i).GetCell(value_idx).StringCellValue;

                    }
                    catch (Exception e)
                    {
                        value = configSheet.GetRow(i).GetCell(value_idx).NumericCellValue.ToString();
                    }
                    value_idx++;
                    //add it to the config dic
                    _configDic.Add(name, value);
                }


            }
        }
        /// <summary>
        /// send a email to the address in config file, using the current outlook account
        /// </summary>
        private static void SendEmail(string subject, string content)
        {
            //set address, subject and body of email
            string emailAddr = _configDic["NotificationEmail"];

            string autoPath = _configDic["AutomationPath"];
            string body = content;

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
            SharePointPage sharePointPage = new SharePointPage(_configDic);
            sharePointPage.DownLoadSharePointDoc();
            //change 1 sheet name from BV & SA to BVSA
            //get path and filename
            string savepath = _configDic["LocalSavePath"];
            string filename = _configDic["SharepointFileName"];
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
            string username = _configDic["ShareFolderUserName"];
            string password = _configDic["ShareFolderPassword"];
            //get local save path and server path
            string localPath = _configDic["LocalSavePath"];
            string serverPath = _configDic["LogSchedulePath"];
            //get filename
            string filename = _configDic["LogScheduleFileName"];
            localPath += "\\";
            //launch a command line to connect to the server
            CmdHelper.ConnectState(serverPath, username, password);

            //copy logistic schedule file
            filename += ".xlsx";
            serverPath += "\\";
            File.Copy(serverPath + filename, localPath + filename, true);
            AddToLog(filename + " download completed");
            //copy Burwood stock transfer file
            serverPath = _configDic["BurwoodPath"] + "\\";
            filename = _configDic["BurwoodFileName"];
            filename = GetNewestFileName(serverPath, filename);
            File.Copy(serverPath + filename, localPath + filename, true);
            AddToLog(filename + " download completed");

            //copy Regents transfer stock file
            serverPath = _configDic["RegentsPath"] + "\\";
            filename = _configDic["RegentsFileName"];
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
                HostName = _configDic["HostName"],
                UserName = _configDic["WinScpUsername"],
                Password = _configDic["WinScpPassword"],
                SshHostKeyFingerprint = _configDic["FingerPrint"],
            };

            using (Session session = new Session())
            {
                //connect
                session.Open(options);

                //upload files
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;

                //get path
                string localPath = _configDic["LocalSavePath"];
                string remotePath = _configDic["RemoteSavePath"];
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

            //process toll documents
            ProcessTollExcels();
            //process meridian documents
            ProcessMeridianExcels();
        }

        private static void ProcessTollExcels()
        {
            //get total toll report numbers
            int totalWaitMilliSecs = 0;
            int totalReportNum = int.Parse(_configDic["TotalTollDocuments"]);
            for (int i = 0; i < totalReportNum; i++)
            {
                //read from report
                string savePath = _configDic["LocalSavePath"];
                string filename = _configDic["TollDocumentName" + (i + 1)];
                savePath += "\\";
                string filepath = savePath + filename;
                //check if the file exists
                string extension = ".xlsx";

                if (!File.Exists(filepath + extension))
                {
                    if (!File.Exists(filepath + ".xls"))
                        throw new Exception(filepath + " is not downloaded");
                }
                int linesToBeDeleted = int.Parse(_configDic["LinesToBeDeleted" + (i + 1)]);

                //use library to read an excel file
                ISheet reportsheet = ExcelHelper.ReadExcel(filepath + extension);

                //do the archive
                ExcelHelper.MoveFileToArchive(savePath, filename, true);
                //save
                ExcelHelper.SaveTo(reportsheet, filepath + ".xlsx", linesToBeDeleted);
                AddToLog(filename + " process completed");

            }


        }

        private static void ProcessMeridianExcels()
        {
            //get total meridian report numbers
            int totalWaitMilliSecs = 0;
            int totalReportNum = int.Parse(_configDic["TotalMeridianDocuments"]);
            for (int i = 0; i < totalReportNum; i++)
            {
                //read from report
                string savePath = _configDic["LocalSavePath"];
                string filename = _configDic["OriginalFileName" + (i + 1)];
                savePath += "\\";
                string filepath = savePath + filename;
                string rename = _configDic["Rename" + (i + 1)];
                string extension = ".xls";
                //check if the file exists
                if (!File.Exists(filepath + extension))
                {
                    if (!File.Exists(filepath + ".xls"))
                        throw new Exception(filepath + " is not downloaded");
                }
                //process the file by string
                ExcelHelper.ProcessInvalidExcel(savePath, filename, rename);
                //save the incorrupted file as xlsx
                OfficeExcel.SaveAs(savePath, rename);
                ExcelHelper.MoveFileToArchive(savePath,rename,false);
                //delete he priginal file
                if (File.Exists(savePath + rename + ".xls"))
                    File.Delete(savePath + rename + ".xls");
                AddToLog(rename + "process completed");
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
            //read data and stores it into a dictionary
            ISheet sheet = ExcelHelper.ReadExcel("Defense Automation Config.xlsx");
            //check if the download folder exists, if not create one
            StoreIntoDic(sheet);
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
                TollLoginPage tollLoginPage = new TollLoginPage(_configDic);
                TollReportDownloadPage tollDownloadPage = tollLoginPage.Login();
                string savepath = _configDic["LocalSavePath"];
                savepath += "\\";
                //download first document
                string filename = _configDic["TollDocumentName1"];
                TollGoodReportPage tollGoodReportPage = tollDownloadPage.DownloadGoodDocument();
                tollGoodReportPage.DownLoadReport(savepath + filename);

                AddToLog(filename + " download completed");
                //download 2nd
                filename = _configDic["TollDocumentName2"];
                tollDownloadPage.GoToReportPage();
                TollShipOrderPage tollShipDetailPage = tollDownloadPage.DownLoadShipOrder();
                tollShipDetailPage.DownLoadReport(savepath + filename);

                AddToLog(filename + " download completed");

                //download the 3rd 
                filename = _configDic["TollDocumentName3"];
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
            MeridianPortalPage meridianPortalPage = new MeridianPortalPage(_configDic);
            MeridianNavigationPage meridianNavigationPage = meridianPortalPage.LaunchMeridian();

            //download files
            //if it fails, retry it
            int retryCount = 3;
            while (true)
            {
                try
                {
                    DownLoadPODetailDoc(meridianNavigationPage);
                    break;
                }
                catch (Exception e)
                {
                    AddToLog("Retry Po Detail Download for " + (3 - retryCount) + " times");
                    //if file exists, delete it
                    string savepath = _configDic["LocalSavePath"];
                    string filename = _configDic["OriginalFileName1"];
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
                    DownLoadAccDetailDoc(meridianNavigationPage);
                    break;
                }
                catch (Exception e)
                {
                    AddToLog("Retry Account Detail Download for " + (3 - retryCount) + " times");
                    //if file exists, delete it
                    string savepath = _configDic["LocalSavePath"];
                    string filename = _configDic["OriginalFileName2"];
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
        private static void DownLoadAccDetailDoc(MeridianNavigationPage meridianNavigationPage)
        {
            //go to account payable entry detail page
            MeridianVariableEntryPage accVariableEntryPage = meridianNavigationPage.GotoAccountDetailEntryPage(_configDic);
            MeridianAccountDetailPage accountDetailPage = accVariableEntryPage.AccountEnter();
            ////click open button select detail
            MeridianPopUpWindow accPopUpWindow = accountDetailPage.OpenPoPUpWindow();
            accPopUpWindow.SelectAccountDetailDoc();
            //get full path
            string filename = _configDic["OriginalFileName2"];
            string savepath = _configDic["LocalSavePath"];
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
        private static void DownLoadPODetailDoc(MeridianNavigationPage meridianNavigationPage)
        {
            //go to PO detail
            MeridianVariableEntryPage POVariableEntryPage = meridianNavigationPage.GotoPoDetailEntryPage(_configDic);
            MeridianAccountDetailPage PoAccountDetailPage = POVariableEntryPage.PODetailEnter();
            //click open button select detail
            MeridianPopUpWindow POPopUpWindow = PoAccountDetailPage.OpenPoPUpWindow();
            POPopUpWindow.SelectPODetailDoc();
            //get full path
            string filename = _configDic["OriginalFileName1"];
            string savepath = _configDic["LocalSavePath"];
            savepath += "\\";
            //download PO Detail Report
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
                string autoPath = _configDic["AutomationPath"];
                autoPath += "\\";

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
            string path = _configDic["LogPath"];
            //create folder
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            //create folder of current date
            path += "\\" + DateTime.Today.ToString("d").Replace("/", "-");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            string filename = path + "\\" + DateTime.Now.ToString().Replace("/", "-").Replace(":", " ") + ".log";
            //write log file
            File.WriteAllText(filename, sb.ToString());
        }
    }
}



