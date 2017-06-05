using System;
using System.Collections.Generic;
using System.Threading;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessObjects;
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
                DeleteAllFiles(configSheet.GetRow(4).GetCell(1).StringCellValue);

                DownLoadDocuments(configSheet);

                //delete several lines at the beginning
                ProcessExcels(configSheet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);

            }


            Exit();





        }

        private static void ProcessExcels(ISheet configSheet)
        {


            //get total report numbers
            try
            {
                int totalReportNum = (int)configSheet.GetRow(5).GetCell(1).NumericCellValue;
                for (int i = 0; i < totalReportNum; i++)
                {
                    //read from report
                    string savePath = configSheet.GetRow(4).GetCell(1).StringCellValue;
                    string filename = configSheet.GetRow(6).GetCell(1 + i).StringCellValue;
                    string filepath = savePath + filename;
                    //wait until file exists
                    while (!File.Exists(filepath + ".xlsx"))
                        Thread.Sleep(500);
                    int linesToBeDeleted = (int)configSheet.GetRow(7).GetCell(1 + i).NumericCellValue;
                    ISheet reportsheet = ExcelHelper.ReadExcel(filepath + ".xlsx");

                    //do the archive
                    MoveFileToArchive(savePath, filename);
                    //save
                    ExcelHelper.SaveTo(reportsheet, filepath + ".xlsx", linesToBeDeleted);
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
            DirectoryInfo di=new DirectoryInfo(path);
            foreach (FileInfo fileInfo in di.GetFiles())
            {
                fileInfo.Delete();
            }
        }

        private static ISheet Intialization()
        {
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
                    if (retryCount <= 0)
                    {
                        Console.WriteLine(e.Message);
                        Exit();
                    }
                    retryCount--;
                }
            }

            return sheet;
        }

        private static void DownLoadDocuments(ISheet configSheet)
        {
            //login
            try
            {
                TollLoginPage tlp = new TollLoginPage(configSheet);
                TollReportDownloadPage trdlp = tlp.Login();
                //download first document
                TollReportPage tgp = trdlp.DownloadGoodsDocument();
                tgp.DownLoadReport();
            }
            catch (NoReportsException reportsException)
            {
                Console.WriteLine("The 'TollTotalDocuments' cell could be emtry");
                Exit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Exit();

            }

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
