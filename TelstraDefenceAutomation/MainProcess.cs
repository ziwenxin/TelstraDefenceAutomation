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
using ICSharpCode.SharpZipLib.Tar;
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
                DeleteAllFiles(configSheet.GetRow(4).GetCell(1).StringCellValue);

                //download excel files
                DownLoadTollDocuments(configSheet);
                DownLoadMeridianDocuments(configSheet);

                //delete several lines at the beginning
                ProcessExcels(configSheet);
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
            int totalReportNum = (int)configSheet.GetRow(5).GetCell(1).NumericCellValue;
            for (int i = 0; i < totalReportNum; i++)
            {
                //read from report
                string savePath = configSheet.GetRow(4).GetCell(1).StringCellValue;
                string filename = configSheet.GetRow(6).GetCell(1 + i).StringCellValue;
                string filepath = savePath + filename;
                //check if the file exists
                string extension = ".xlsx";

                if (!File.Exists(filepath + extension))
                {
                    if (!File.Exists(filepath + ".xls"))
                        throw new Exception(filepath + " is not downloaded");
                }
                int linesToBeDeleted = (int)configSheet.GetRow(7).GetCell(1 + i).NumericCellValue;

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
                    //process the file by string
                    ExcelHelper.ProcessInvalidExcel(savePath, filename);
                }
                Console.WriteLine(filename + " process completed");
            }



        }

        //delete all files but not folders in a folder
        private static void DeleteAllFiles(string path)
        {
            Console.WriteLine("Deleting all previous files...");
            DirectoryInfo di = new DirectoryInfo(path);
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
                    if (WebDriver.ChromeDriver != null)
                        WebDriver.ChromeDriver.Quit();
                    if (retryCount <= 0)
                    {
                        throw e;
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
                string savepath = configSheet.GetRow(4).GetCell(1).StringCellValue;
                //download first document
                string filename = configSheet.GetRow(6).GetCell(1).StringCellValue;
                TollGoodReportPage tollGoodReportPage = tollDownloadPage.DownloadGoodDocument();
                tollGoodReportPage.DownLoadReport(savepath + filename);
                Console.WriteLine(filename + " download completed");
                //download 2nd
                filename = configSheet.GetRow(6).GetCell(2).StringCellValue;
                tollDownloadPage.GoToReportPage();
                TollShipOrderPage tollShipDetailPage = tollDownloadPage.DownLoadShipOrder();
                tollShipDetailPage.DownLoadReport(savepath + filename);

                Console.WriteLine(filename + "download completed");

                //download the 3rd 
                filename = configSheet.GetRow(6).GetCell(3).StringCellValue;
                tollDownloadPage.GoToReportPage();
                TollSOHDetailPage tollSohDetailPage = tollDownloadPage.DownloadSOHDetail();
                tollSohDetailPage.DownLoadReport(savepath + filename);
                Console.WriteLine(filename + "download completed");

            }
            catch (NoReportsException reportsException)
            {
                Console.WriteLine("The 'TollTotalDocuments' cell could not be emtry");
                throw reportsException;
            }
            catch (Exception e)
            {
                throw e;

            }

        }

        private static void DownLoadMeridianDocuments(ISheet configSheet)
        {
            Console.WriteLine("Downloading Meridian documents...");
            //go to the portal of meridian
            MeridianPortalPage meridianPortalPage = new MeridianPortalPage(configSheet);
            MeridianNavigationPage meridianNavigationPage = meridianPortalPage.LaunchMeridian();

            //download files
            DownLoadPODetailDoc(configSheet, meridianNavigationPage);
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
            //get full path
            string filename = configSheet.GetRow(6).GetCell(5).StringCellValue;
            string savepath = configSheet.GetRow(4).GetCell(1).StringCellValue;
            ////download PO Detail Reprrt
            accountDetailPage.DownLoadAccountDetailDoc(savepath + filename);
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
            //get full path
            string filename = configSheet.GetRow(6).GetCell(4).StringCellValue;
            string savepath = configSheet.GetRow(4).GetCell(1).StringCellValue;
            //download PO Detail Reprrt
            PoAccountDetailPage.DownLoadPoDetailDoc(savepath + filename);
            Console.WriteLine(filename + " download completed");
        }

        public static void RescheduleTask()
        {
            
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
