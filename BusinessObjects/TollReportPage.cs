using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects
{
    public class TollReportPage
    {
        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_txtValue")]
        public IWebElement FromDateField { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl05_txtValue")]
        public IWebElement ToDateField { get; set; }


        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_txtValue")]
        public IWebElement OwnerIdCbl { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_divDropDown_ctl08")]
        public IWebElement OwnerIdCb { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl00")]
        public IWebElement ViewReportBtn { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl05_ctl04_ctl00_ButtonImg")]
        public IWebElement SaveIcon { get; set; }

        [FindsBy(How = How.XPath,Using = "//a[@title='Excel']")]
        public IWebElement ExcelSaveLink { get; set; }


        public TollReportPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }

        public void AddFilter()
        {
            //the data range should be from the first date of this year to today
            int thisYear = DateTime.Now.Year;
            FromDateField.SendKeys(new DateTime(thisYear, 1, 1).ToString());
            ToDateField.SendKeys(DateTime.Today.ToString());
            OwnerIdCbl.Click();
            OwnerIdCb.Click();


        }

        public void GenerateReport()
        {
            AddFilter();
            ViewReportBtn.Click();
            
        }

        public void DownLoadReport()
        {
            IWebDriver driver = WebDriver.ChromeDriver;
            GenerateReport();
            //wait until the save icon exists
            WebDriverWait wait=new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ReportViewer1_ctl05_ctl04_ctl00_ButtonImg")));
            Thread.Sleep(1000);
            SaveIcon.Click();
            ExcelSaveLink.Click();
        }


    }
}
