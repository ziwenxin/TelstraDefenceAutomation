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

        public List<IWebElement> ReportLinks { get; set; }

        [FindsBy(How = How.XPath, Using = "//iframe")]
        public IWebElement ReportFrame { get; set; }

        public TollReportDownloadPage(ISheet Configsheet)
        {
            //get configsheet and driver
            this.configSheet = Configsheet;
            IWebDriver driver = WebDriver.ChromeDriver;
            PageFactory.InitElements(driver, this);

            //get total report numbers
            int totalReportNum = (int)configSheet.GetRow(5).GetCell(1).NumericCellValue;
            if (totalReportNum <= 0)
                throw new NoReportsException();
            //swtich to report frame
            GoToReportPage();
            WebDriver.ChromeDriver.SwitchTo().Frame(ReportFrame);
            ReportLinks = new List<IWebElement>();
            //get reports names and find the link
            for (int i = 0; i < totalReportNum; i++)
            {
                string reportName = configSheet.GetRow(6).GetCell(1 + i).StringCellValue;
                IWebElement webElement = driver.FindElement(By.XPath("//a[text()='" + reportName + "']"));
                ReportLinks.Add(webElement);

            }

        }

        public void GoToReportPage()
        {
            WebDriver.ChromeDriver.Navigate().GoToUrl(configSheet.GetRow(3).GetCell(1).StringCellValue);


        }

        public TollReportPage DownloadGoodsDocument()
        {
            //click the link to download
            for (int i = 0; i < ReportLinks.Count; i++)
            {
                if (i != 0)
                    GoToReportPage();
                ReportLinks[i].Click();
            }
            return new TollReportPage();
        }
    }

}
