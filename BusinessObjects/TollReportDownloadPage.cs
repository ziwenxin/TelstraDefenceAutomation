using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Exceptions;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects
{
    public class TollReportDownloadPage
    {
        private ISheet configSheet;

        public IWebElement GoodReportLink { get; set; }

        public IWebElement ShipDetailLink { get; set; }

        public IWebElement SOHDetailLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//iframe")]
        public IWebElement ReportFrame { get; set; }

        public TollReportDownloadPage(ISheet Configsheet)
        {
            //get configsheet and driver
            this.configSheet = Configsheet;
            IWebDriver driver = WebDriver.ChromeDriver;
            PageFactory.InitElements(driver, this);

            //find elements with the file names from config file
            int totalDocuments = (int) configSheet.GetRow(5).GetCell(1).NumericCellValue;
            if (totalDocuments <= 0)
                throw new NoReportsException();
            //swtich to report frame
            GoToReportPage();

            //set links
            GoodReportLink = driver.FindElement(By.XPath("//a[text()='"+configSheet.GetRow(6).GetCell(1).StringCellValue+"']"));
            ShipDetailLink = driver.FindElement(By.XPath("//a[text()='" + configSheet.GetRow(6).GetCell(2).StringCellValue + "']"));
            SOHDetailLink = driver.FindElement(By.XPath("//a[text()='" + configSheet.GetRow(6).GetCell(3).StringCellValue + "']"));



        }

        public void GoToReportPage()
        {
            WebDriver.ChromeDriver.Navigate().GoToUrl(configSheet.GetRow(3).GetCell(1).StringCellValue);
            WebDriver.ChromeDriver.SwitchTo().Frame(ReportFrame);

        }

        public TollGoodReportPage DownloadGoodDocument()
        {
            //go to report page and click the link

            GoodReportLink.Click();
            return new TollGoodReportPage();
        }

        public TollShipDetailPage DownLoadShipDetail()
        {
            //go to report page and click the link

            GoodReportLink.Click();
            return new TollShipDetailPage();
        }

        public TollSOHDetailPage DownloadSOHDetail()
        {
            //go to report page and click the link

            SOHDetailLink.Click();
            return new TollSOHDetailPage();
        }
    }

}
