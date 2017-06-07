using System;
using System.Collections.Generic;
using System.Threading;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessObjects;
using BusinessObjects.MERIDIAN;
using Common;
using Exceptions;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using PropertyCollection;

namespace TelstraDefenceAutomation
{
    public class MainProcess
    {
        static void Main(string[] args)
        {
            try
            {
                //read settings and set default download folder for chrome
                ISheet configSheet = Intialization();

                //before automation, delete all files in the save folder
                //DeleteAllFiles(configSheet.GetRow(4).GetCell(1).StringCellValue);

                DownLoadMeridianDocuments(configSheet);
                //DownLoadTollDocuments(configSheet);

                //delete several lines at the beginning
                //ProcessExcels(configSheet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine("\r\n Press Any Key To Exit");
                Console.ReadKey();
            }


            Exit();





        }

        private static void ProcessExcels(ISheet configSheet)
        {
            Console.WriteLine("Processing Excel files...");
            int totalWaitMilliSecs = 0;
            //get total report numbers
            try
            {
                int totalReportNum = (int)configSheet.GetRow(5).GetCell(1).NumericCellValue;
                for (int i = 3; i < totalReportNum; i++)
                {
                    //read from report
                    string savePath = configSheet.GetRow(4).GetCell(1).StringCellValue;
                    string filename = configSheet.GetRow(6).GetCell(1 + i).StringCellValue;
                    string filepath = savePath + filename;
                    //wait until file exists
                    string extension = ".xlsx";
                    while (!File.Exists(filepath + extension))
                    {
                        if (File.Exists(filepath + ".xls"))
                        {
                            extension = ".xls";
                            break;
                        }
                        Thread.Sleep(500);
                        totalWaitMilliSecs += 500;
                        if(totalWaitMilliSecs>20000)
                            throw new Exception( filename+" is not downloaded");
                    }
                    int linesToBeDeleted = (int)configSheet.GetRow(7).GetCell(1 + i).NumericCellValue;
                    ISheet reportsheet = ExcelHelper.ReadExcel(filepath + extension);

                    //do the archive
                    MoveFileToArchive(savePath, filename);
                    //save
                    ExcelHelper.SaveTo(reportsheet, filepath + ".xlsx", linesToBeDeleted);
                    Console.WriteLine(filename+" process completed");
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Exit();
            }


        }

        //delete all files but not folders in a folder
        private static void DeleteAllFiles(string path)
        {
            Console.WriteLine("Deleting all previous files...");
            DirectoryInfo di=new DirectoryInfo(path);
            foreach (FileInfo fileInfo in di.GetFiles())
            {
                fileInfo.Delete();
            }
            Console.WriteLine("Delete completed");
        }

        private static ISheet Intialization()
        {
            Console.WriteLine("Inialising...");

            int retryCount = 3;
            //read data
            ISheet sheet = ExcelHelper.ReadExcel("Config.xlsx");

            //check if the download folder exists, if not create one
            string path = sheet.GetRow(4).GetCell(1).StringCellValue;
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
                    if(WebDriver.ChromeDriver!=null)
                        WebDriver.ChromeDriver.Quit();
                    if (retryCount <= 0)
                    {
                        Console.WriteLine(e.Message);
                        Exit();
                    }
                    retryCount--;
                }
            }
            Console.WriteLine("Initialization completed");
            return sheet;
        }

        private static void DownLoadTollDocuments(ISheet configSheet)
        {
            Console.WriteLine("DownLoading documents from Toll...");
            //login
            try
            {
                TollLoginPage tollLoginPage = new TollLoginPage(configSheet);
                TollReportDownloadPage tollDownloadPage = tollLoginPage.Login();
                //download first document
                TollGoodReportPage tollGoodReportPage = tollDownloadPage.DownloadGoodDocument();
                tollGoodReportPage.DownLoadReport();
                Console.WriteLine("TelDef - Goods Receipt By Date Range download completed");
                //download 2nd
                tollDownloadPage.GoToReportPage();
                TollShipOrderPage tollShipDetailPage = tollDownloadPage.DownLoadShipOrder();
                tollShipDetailPage.DownLoadReport();
                Console.WriteLine("TelDef - Shipped Order Report download completed");

                //download the 3rd 
                tollDownloadPage.GoToReportPage();
                TollSOHDetailPage tollSohDetailPage = tollDownloadPage.DownloadSOHDetail();
                tollSohDetailPage.DownLoadReport();
                Console.WriteLine("TelDef - SOH Detail Report download completed");

            }
            catch (NoReportsException reportsException)
            {
                Console.WriteLine("The 'TollTotalDocuments' cell could not be emtry");
                Exit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Exit();

            }

        }

        private static void DownLoadMeridianDocuments(ISheet configSheet)
        {
            Console.WriteLine("Downloading Meridian documents...");
            //go to the portal of meridian
            MeridianPortalPage meridianPortalPage = new MeridianPortalPage(configSheet);
            MeridianNavigationPage meridianNavigationPage = meridianPortalPage.LaunchMeridian();
            //DownLoadPODetailDoc(configSheet, meridianNavigationPage);

            DownLoadAccDetailDoc(configSheet, meridianNavigationPage);
        }

        private static void DownLoadAccDetailDoc(ISheet configSheet, MeridianNavigationPage meridianNavigationPage)
        {
            //go to account payable entry detail page
            MeridianVariableEntryPage accVariableEntryPage = meridianNavigationPage.GotoAccountDetailEntryPage(configSheet);
            MeridianAccountDetailPage accountDetailPage = accVariableEntryPage.AccountEnter();
            ////click open button select detail
            MeridianPopUpWindow accPopUpWindow = accountDetailPage.OpenPoPUpWindow();
            accPopUpWindow.SelectAccountDetailDoc();
            ////download PO Detail Reprrt
            accountDetailPage.DownLoadAccountDetailDoc();
            Console.WriteLine("Account Detail download completed");
        }

        private static void DownLoadPODetailDoc(ISheet configSheet, MeridianNavigationPage meridianNavigationPage)
        {
            //go to PO detail
            MeridianVariableEntryPage POVariableEntryPage = meridianNavigationPage.GotoPoDetailEntryPage(configSheet);
            MeridianAccountDetailPage PoAccountDetailPage = POVariableEntryPage.PODetailEnter();
            //click open button select detail
            MeridianPopUpWindow POPopUpWindow = PoAccountDetailPage.OpenPoPUpWindow();
            POPopUpWindow.SelectPODetailDoc();
            //download PO Detail Reprrt
            PoAccountDetailPage.DownLoadPoDetailDoc();
            Console.WriteLine("Po Detail download completed");
        }

        private static void MoveFileToArchive(string savePath, string filename)
        {
            //save set archivepath and archive file name
            string archivePath = savePath + "Archive/";
            if (!Directory.Exists(archivePath))
                Directory.CreateDirectory(archivePath);
            //set date format
            string dateStr = DateTime.Today.ToString("d");
            dateStr = dateStr.Replace("/", "-");
            //set a data folder in the archive folder
            archivePath += dateStr + "/";
            if (!Directory.Exists(archivePath))
                Directory.CreateDirectory(archivePath);
            string archiveFilename = filename + " " + dateStr;
            //set destination path and original path
            string OriginalPath = savePath + filename + ".xlsx";
            string dstPath = archivePath + archiveFilename + ".xlsx";
            //if the archive file exists, delete it
            if (File.Exists(dstPath))
                File.Delete(dstPath);
            //copy the file to archive folder
            File.Copy(OriginalPath, dstPath);
            //delete the original file
            File.Delete(OriginalPath);
        }

        private static void Exit()
        {
            Console.WriteLine("The automation will be closed in 5 secs");
            //close the automation
            Thread.Sleep(5000);
            WebDriver.ChromeDriver.Quit();
            Environment.Exit(0);
        }




    }
}
