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
    public class TollGoodReportPage :TollReportPage
    {
        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_txtValue")]
        public IWebElement FromDateField { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl05_txtValue")]
        public IWebElement ToDateField { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_txtValue")]
        public IWebElement OwnerIdCbl { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_divDropDown_ctl08")]
        public IWebElement OwnerIdCb { get; set; }




        public TollGoodReportPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }

        public override void AddFilter()
        {

            //the data range should be from the first date of this year to today
            int thisYear = DateTime.Now.Year;
            FromDateField.SendKeys(new DateTime(thisYear, 1, 1).ToString());
            ToDateField.SendKeys(DateTime.Today.ToString());
            //check the owner id
            OwnerIdCbl.Click();
            OwnerIdCb.Click();

        }

       


        public override void DownLoadReport()
        {
            //add filter and generate the report
            AddFilter();
            ViewReportBtn.Click();

            IWebDriver driver = WebDriver.ChromeDriver;
            //wait until the loading finish
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.Id("ReportViewer1_AsyncWait_Wait")));
            Thread.Sleep(1000);
            SaveIcon.Click();
            ExcelSaveLink.Click();
        }


    }
}
