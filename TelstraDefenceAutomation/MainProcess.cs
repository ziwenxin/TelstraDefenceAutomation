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
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;
using Exception = System.Exception;


namespace TelstraDefenceAutomation
{
    public class MainProcess
    {

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
                retryTimes = int.Parse(ConfigHelper._configDic["RerunTimes"]);

                //delete too old archives
                FileHelper.DeleteOldArchive(ConfigHelper._configDic["LocalSavePath"] + "\\Archive");



                //download the supplier documents
                OutlookHelper.DownloadAttachments();

                //delete the first row and sheet1 for secure edge and delete 4 rows for avnet
                ProcessSalesExcels();

                //download excel files
                DownLoadTollDocuments();
                DownLoadMeridianDocuments();

                //delete several lines at the beginning
                ProcessWebExcels();

                //copy files from share folder
                DownLoadShareFolderDocs();

                //download files from share point
                DownLoadSharePointDoc();


                //upload to server
                WinScpHelper.UploadFiles();


                //delete all the files in the save folder, if any
                FileHelper.DeleteAllFiles(ConfigHelper._configDic["LocalSavePath"]);


                //run lavastorm program
                if (ConfigHelper._configDic["EnableAutomation?"].ToLower() == "yes")
                {

                    CmdHelper.RunLavaStorm();

                }

                //renew retry times

                retryTimes = 3;

                #region SucessEmail

                //send email
                string subject = "Automation Success";
                string content = "Hello," + Environment.NewLine + Environment.NewLine + "The bot has completed a successful run today at " + DateTime.Now + Environment.NewLine + Environment.NewLine + "Kind Regards," + Environment.NewLine + "The Defence Inventory Data Hub bot";
                OutlookHelper.SendEmail(subject, content);

                #endregion
            }
            catch (Exception e)
            {
                //log
                LogHelper.AddToLog(e.ToString());
                //if still needs to retry
                if (retryTimes > 0)
                {
                    //reschedule one run
                    try
                    {
                        RescheduleTask();
                        retryTimes--;
                    }
                    catch (Exception exception)
                    {
                        LogHelper.AddToLog(exception.Message);
                    }
                }
                //notify admin
                else
                {

                    //reset retry times
                    retryTimes = 3;
                    try
                    {
                        #region FailureEmail

                        //set content and subject
                        string autoPath = ConfigHelper._configDic["AutomationPath"];
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
                                         Environment.NewLine + Environment.NewLine +
                                         "Thank you for helping me complete my run," + Environment.NewLine +
                                         "The Defence Inventory Data Hub bot";
                        OutlookHelper.SendEmail(subject, content);

                        #endregion

                    }
                    catch (Exception exception)
                    {
                        LogHelper.AddToLog(exception.Message);
                    }
                }



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
                    LogHelper.AddToLog(e.Message);
                }
            }





            Exit();
            #endregion





        }

        /// <summary>
        /// process excels from sales order
        /// </summary>
        private static void ProcessSalesExcels()
        {
            //process the two files
            ExcelProcesser.ProcessAvnetExcel();
            ExcelProcesser.ProcessSucureExcel();
            //make all the file in upper case, which indicates that they are all processed
            string saleFolder = ConfigHelper._configDic["LocalSavePath"] + "\\SalesOrderHistory\\";
            FileHelper.AddTimeStamps(saleFolder);
        }


        /// <summary>
        /// it will read data from config sheet to a dictionary
        /// </summary>
        /// <param name="configSheet">the configuration sheet</param>
        private static void StoreIntoDic(ISheet configSheet)
        {
            //initial
            ConfigHelper._configDic = new Dictionary<string, string>();
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
                    ConfigHelper._configDic.Add(name, value);
                }


            }
        }


        /// <summary>
        /// download 'Deployment Planning and Tracking' from share point
        /// </summary>
        private static void DownLoadSharePointDoc()
        {
            LogHelper.AddToLog("Downloading from share point...");
            //get path and filename
            string savepath = ConfigHelper._configDic["LocalSavePath"];
            string filename = ConfigHelper._configDic["SharepointFileName"];
            savepath += "\\";
            //download file from share point, if not exists
            if (!File.Exists(savepath + filename + ".xlsx"))
            {
                SharePointPage sharePointPage = new SharePointPage();
                sharePointPage.DownLoadSharePointDoc();
                //change 1 sheet name from BV & SA to BVSA

                //set sheet name
                OfficeExcelHelper.ChangeSheetName(savepath, filename, "BV & SA", "BVSA");
            }



            LogHelper.AddToLog("DownLoad from share point completed");
        }

        /// <summary>
        /// download 'Logistics','All-CECs-StockTransfer Burwood' and 'All-CECs-StockTransfer-Regents' from share folder
        /// </summary>
        private static void DownLoadShareFolderDocs()
        {
            LogHelper.AddToLog("Downloading files from share folder...");

            //get username and password
            string username = ConfigHelper._configDic["ShareFolderUserName"];
            string password = ConfigHelper._configDic["ShareFolderPassword"];
            //get local save path and server path
            string localPath = ConfigHelper._configDic["LocalSavePath"];
            string serverPath = ConfigHelper._configDic["LogSchedulePath"];
            //get filename
            string filename = ConfigHelper._configDic["LogScheduleFileName"];
            localPath += "\\";
            //launch a command line to connect to the server
            CmdHelper.ConnectState(serverPath, username, password);

            //copy logistic schedule file
            filename += ".xlsx";
            serverPath += "\\";
            File.Copy(serverPath + filename, localPath + filename, true);
            LogHelper.AddToLog(filename + " download completed");
            //copy Burwood stock transfer file
            serverPath = ConfigHelper._configDic["BurwoodPath"] + "\\";
            filename = ConfigHelper._configDic["BurwoodFileName"];
            filename = FileHelper.GetNewestFileName(serverPath, filename);
            File.Copy(serverPath + filename, localPath + filename, true);
            LogHelper.AddToLog(filename + " download completed");

            //copy Regents transfer stock file
            serverPath = ConfigHelper._configDic["RegentsPath"] + "\\";
            filename = ConfigHelper._configDic["RegentsFileName"];
            filename = FileHelper.GetNewestFileName(serverPath, filename);
            File.Copy(serverPath + filename, localPath + filename, true);
            LogHelper.AddToLog(filename + " download completed");
        }



        /// <summary>
        /// process the excel files downloaded from 'Toll' and 'Meridian', it mainly delete several lines from the top of the documents
        /// </summary>
        private static void ProcessWebExcels()
        {
            LogHelper.AddToLog("Processing Excel files...");

            //process toll documents
            ExcelProcesser.ProcessTollExcels();
            //process meridian documents
            ExcelProcesser.ProcessMeridianExcels();

        }


        /// <summary>
        /// initial the webdriver and read data from config file
        /// </summary>
        /// <returns>the work sheet of config file</returns>
        private static ISheet Intialization()
        {
            //log
            LogHelper.AddToLog("Inialising...");

            int retryCount = 3;
            //read data and stores it into a dictionary
            ISheet sheet = ExcelHelper.ReadExcel("Defense Automation Config.xlsx");
            //check if the download folder exists, if not create one
            StoreIntoDic(sheet);
            string path = ConfigHelper._configDic["LocalSavePath"];
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            if (!Directory.Exists(path + "\\Archive"))
                Directory.CreateDirectory(path + "\\Archive");
            if (!Directory.Exists(path + "\\SalesOrderHistory"))
                Directory.CreateDirectory(path + "\\SalesOrderHistory");
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
            LogHelper.AddToLog("Initialization completed");
            return sheet;
        }

        /// <summary>
        /// download 'TelDef - Shipped Order Report','TelDef - SOH Detail' and 'TelDef - Goods Receipt By Date Range' from toll 
        /// </summary>
        private static void DownLoadTollDocuments()
        {
            LogHelper.AddToLog("DownLoading documents from Toll...");

            try
            {
            //login
                TollLoginPage tollLoginPage = new TollLoginPage();
                TollReportDownloadPage tollDownloadPage = tollLoginPage.Login();
                string savepath = ConfigHelper._configDic["LocalSavePath"];
                savepath += "\\";
                //download first document
                string filename = ConfigHelper._configDic["TollDocumentName1"];
                //download 1st if not exists
                if (!File.Exists(savepath + filename+" v2.xlsx")&&!File.Exists(savepath + FileHelper.RemoveV2(filename) + ".xlsx"))
                {
                    TollGoodReportPage tollGoodReportPage = tollDownloadPage.DownloadGoodDocument();
                    tollGoodReportPage.DownLoadReport(savepath + filename);
                //rename the file with v2
                File.Move(savepath+filename+".xlsx",savepath+filename+" v2.xlsx");
                    LogHelper.AddToLog(filename + " download completed");
                }

                //download 2nd
                filename = ConfigHelper._configDic["TollDocumentName2"];
                //download 2nd if not exists
                if (!File.Exists(savepath + filename+".xlsx") && !File.Exists(savepath + FileHelper.RemoveV2(filename) + ".xlsx"))
                {
                    tollDownloadPage.GoToReportPage();
                    TollShipOrderPage tollShipDetailPage = tollDownloadPage.DownLoadShipOrder();
                    tollShipDetailPage.DownLoadReport(savepath + filename);

                    LogHelper.AddToLog(filename + " download completed");
                }

                //download the 3rd 
                filename = ConfigHelper._configDic["TollDocumentName3"];
                //download if not exists
                if (!File.Exists(savepath + filename+ ".xlsx") && !File.Exists(savepath + FileHelper.RemoveV2(filename) + ".xlsx"))
                {
                    tollDownloadPage.GoToReportPage();
                    TollSOHDetailPage tollSohDetailPage = tollDownloadPage.DownloadSOHDetail();
                    tollSohDetailPage.DownLoadReport(savepath + filename);

                    LogHelper.AddToLog(filename + " download completed");
                }
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
            LogHelper.AddToLog("Downloading Meridian documents...");
            //go to the portal of meridian
            MeridianPortalPage meridianPortalPage = new MeridianPortalPage();
            MeridianNavigationPage meridianNavigationPage = meridianPortalPage.LaunchMeridian();

            //get filename and savepath
            string savepath = ConfigHelper._configDic["LocalSavePath"];
            string filename = ConfigHelper._configDic["OriginalFileName1"];
            savepath += "\\";
            string rename = ConfigHelper._configDic["Rename1"];
            //download files
            //if it fails, retry it
            int retryCount = 3;
            while (true)
            {
                try
                {
                    //download if not exists
                    if (!File.Exists(savepath+filename+".xls")&& !File.Exists(savepath + rename+".xlsx"))
                    DownLoadPODetailDoc(meridianNavigationPage);
                    break;
                }
                catch (Exception e)
                {
                    LogHelper.AddToLog("Retry Po Detail Download for " + (4 - retryCount) + " times");
                    //if file exists, delete it
  
                    FileHelper.DeleteFile(savepath, filename);
                    if (retryCount <= 0)
                        throw e;
                    retryCount--;
                    //switch back to main frame
                    WebDriver.ChromeDriver.SwitchTo().DefaultContent();
                }
            }
            //retry
            retryCount = 3;

            filename = ConfigHelper._configDic["OriginalFileName2"];
            rename= ConfigHelper._configDic["Rename2"];
            while (true)
            {
                try
                {
                    //download if not exists
                    if (!File.Exists(savepath + filename + ".xls") && !File.Exists(savepath + rename + ".xlsx"))
                        DownLoadAccDetailDoc(meridianNavigationPage);
                    break;
                }
                catch (Exception e)
                {
                    LogHelper.AddToLog("Retry Account Detail Download for " + (4 - retryCount) + " times");
                    //if file exists, delete it

                    if (retryCount <= 0)
                        throw e;
                    retryCount--;
                    //switch back to main frame
                    WebDriver.ChromeDriver.SwitchTo().DefaultContent();
                }
            }
            //log
            LogHelper.AddToLog("Download Meridian Documents completed.");
        }

        /// <summary>
        /// download 'Accounting_Details_from_meridian'
        /// </summary>
        /// <param name="meridianNavigationPage"></param>
        private static void DownLoadAccDetailDoc(MeridianNavigationPage meridianNavigationPage)
        {
            //go to account payable entry detail page
            MeridianVariableEntryPage accVariableEntryPage = meridianNavigationPage.GotoAccountDetailEntryPage();
            MeridianAccountDetailPage accountDetailPage = accVariableEntryPage.AccountEnter();
            ////click open button select detail
            MeridianPopUpWindow accPopUpWindow = accountDetailPage.OpenPoPUpWindow();
            accPopUpWindow.SelectAccountDetailDoc();
            //get full path
            string filename = ConfigHelper._configDic["OriginalFileName2"];
            string savepath = ConfigHelper._configDic["LocalSavePath"];
            savepath += "\\";

            ////download PO Detail Reprrt
            accountDetailPage.DownLoadAccountDetailDoc(savepath + filename);

            LogHelper.AddToLog(filename + " download completed");
        }

        /// <summary>
        /// download 'PO_DETAILS_REPORT'
        /// </summary>
        /// <param name="meridianNavigationPage"></param>
        private static void DownLoadPODetailDoc(MeridianNavigationPage meridianNavigationPage)
        {
            //go to PO detail
            MeridianVariableEntryPage POVariableEntryPage = meridianNavigationPage.GotoPoDetailEntryPage();
            MeridianAccountDetailPage PoAccountDetailPage = POVariableEntryPage.PODetailEnter();
            //click open button select detail
            MeridianPopUpWindow POPopUpWindow = PoAccountDetailPage.OpenPoPUpWindow();
            POPopUpWindow.SelectPODetailDoc();
            //get full path
            string filename = ConfigHelper._configDic["OriginalFileName1"];
            string savepath = ConfigHelper._configDic["LocalSavePath"];
            savepath += "\\";
            //download PO Detail Report
            PoAccountDetailPage.DownLoadPoDetailDoc(savepath + filename);

            LogHelper.AddToLog(filename + " download completed");

        }

        /// <summary>
        /// reschedule the program after 1 min
        /// </summary>
        private static void RescheduleTask()
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
                string autoPath = ConfigHelper._configDic["AutomationPath"];
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
            LogHelper.AddToLog("The automation will be closed in 5 secs");
            //close the automation
            Thread.Sleep(5000);
            if (WebDriver.ChromeDriver != null)
                WebDriver.ChromeDriver.Quit();
            Environment.Exit(0);
        }




        /// <summary>
        /// write all the runtime information to a log file
        /// </summary>
        private static void WriteLogFile()
        {
            //get path and file name
            string path = ConfigHelper._configDic["LogPath"];
            //create folder
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            //create folder of current date
            path += "\\" + DateTime.Today.ToString("d").Replace("/", "-");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            //define filename
            string filename = path + "\\" + DateTime.Now.ToString().Replace("/", "-").Replace(":", " ") + ".log";
            //write
            LogHelper.WriteLog(filename);
        }
    }
}



