using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects
{
    public class TollReportPage
    {
        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl00")]
        public IWebElement ViewReportBtn { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl05_ctl04_ctl00_ButtonImg")]
        public IWebElement SaveIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Excel']")]
        public IWebElement ExcelSaveLink { get; set; }

        public virtual void AddFilter() { }

        public void DownLoadReport()
        {
            IWebDriver driver = WebDriver.ChromeDriver;
            //wait for the page loaded
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ReportViewer1_ctl04_ctl00")));
            //add filter and generate the report
            AddFilter();
            ViewReportBtn.Click();

            //waut until the loading appears
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ReportViewer1_AsyncWait_Wait")));
            //wait until the loading finish
            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.Id("ReportViewer1_AsyncWait_Wait")));
            SaveIcon.Click();
            ExcelSaveLink.Click();
        }

    }
}


