using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessObjects;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace TelstraDefenceAutomation
{
    public class MainProcess
    {
        static void Main(string[] args)
        {
            ReadData();

            TollLoginPage tw = new TollLoginPage();
            tw.Login();
            tw.DownloadGoodsDocument();
        }

        private static void ReadData()
        {
            //get toll data
            try
            {
                using (FileStream fs = new FileStream("TollData.xlsx", FileMode.Open, FileAccess.Read))
                {
                    XSSFWorkbook hssfWb = new XSSFWorkbook(fs);
                    sheet = hssfWb.GetSheet("Sheet1");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Failed to read the data from the data file");
                Environment.Exit(0);
            }
        }
    }
}
