using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects
{
    public class TollReportPage
    {
        private ISheet sheet;

        [FindsBy(How = How.XPath, Using = "//a[text()='TelDef - Goods Receipt By Date Range']")]
        public IWebElement GoodReportLink { get; set; }

        [FindsBy(How = How.XPath,Using = "//iframe")]
        public IWebElement ReportFrame { get; set; }

        public TollReportPage(ISheet sheet)
        {
            this.sheet = sheet;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }

        public void GoToReportPage()
        {
            WebDriver.ChromeDriver.Navigate().GoToUrl(sheet.GetRow(1).GetCell(3).StringCellValue);
            //swtich to report frame
            WebDriver.ChromeDriver.SwitchTo().Frame(ReportFrame);

        }

        public void DownloadGoodsDocument()
        {
            GoToReportPage();
            GoodReportLink.Click();

        }
    }
}
