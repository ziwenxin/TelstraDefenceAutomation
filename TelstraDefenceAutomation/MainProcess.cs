using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessObjects;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace TelstraDefenceAutomation
{
    public class MainProcess
    {
        static void Main(string[] args)
        {
            ISheet sheet = ReadData();

            TollLoginPage tlp = new TollLoginPage(sheet);
            TollReportPage trp= tlp.Login();

            trp.DownloadGoodsDocument();
        }

        private static ISheet ReadData()
        {
            //get toll data
            try
            {
                using (FileStream fs = new FileStream("TollData.xlsx", FileMode.Open, FileAccess.Read))
                {
                    XSSFWorkbook hssfWb = new XSSFWorkbook(fs);
                    ISheet sheet = hssfWb.GetSheet("Sheet1");
                    return sheet;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Failed to read the data from the data file");
                Environment.Exit(0);
            }
            return null;
        }
    }
}
