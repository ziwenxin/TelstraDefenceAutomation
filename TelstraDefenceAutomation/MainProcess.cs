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
            //read settings and set default download folder for chrome
            ISheet configsheet = Intialization();

            //DownLoadDocuments(configsheet);

            //delete several lines at the beginning
            ProcessExcels(configsheet);

            Exit();





        }

        private static void ProcessExcels(ISheet configsheet)
        {
            //get total report numbers
            try
            {
                int totalReportNum = (int)configsheet.GetRow(5).GetCell(1).NumericCellValue;
                for (int i = 0; i < totalReportNum; i++)
                {
                    //read from report
                    string savePath = configsheet.GetRow(4).GetCell(1).StringCellValue;
                    string filename = configsheet.GetRow(6).GetCell(1 + i).StringCellValue;
                    string filepath = savePath + filename;
                    int linesToBeDeleted = (int)configsheet.GetRow(7).GetCell(1 + i).NumericCellValue;
                    ISheet reportsheet = ExcelHelper.ReadExcel(filepath + ".xlsx");
                    //delete the several lines
                    for (int j = 0; j < linesToBeDeleted; j++)
                    {
                        IRow row = reportsheet.GetRow(j);
                        reportsheet.RemoveRow(row);
                    }
                    //move the rows below to the top

                    //save
                    ExcelHelper.SaveTo(reportsheet, filepath + ".xlsx");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Exit();
            }


        }

        private static ISheet Intialization()
        {
            //read data
            ISheet sheet = ExcelHelper.ReadExcel("TollData.xlsx");

            //check if the download folder exists, if not create one
            string path = sheet.GetRow(4).GetCell(1).StringCellValue;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            //change default download location
            WebDriver.Init(path);
            return sheet;
        }

        private static void DownLoadDocuments(ISheet configsheet)
        {
            //login
            try
            {
                TollLoginPage tlp = new TollLoginPage(configsheet);
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

        private static void MoveFileToArchive()
        {
            
        }

        private static void Exit()
        {
            Console.WriteLine("The automation will be closed in 5 secs");
            //close the automation
            Thread.Sleep(5000);
            WebDriver.ChromeDriver.Quit();
        }




    }
}
