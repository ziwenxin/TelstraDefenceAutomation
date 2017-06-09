using System;
using System.Collections.Generic;
using System.IO;
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

        public void DownLoadReport(string fullpath)
        {
            IWebDriver driver = WebDriver.ChromeDriver;
            //wait for the page loaded
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ReportViewer1_ctl04_ctl00")));
            //add filter and generate the report
            AddFilter();
            ViewReportBtn.Click();

            //waut until the loading appears
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ReportViewer1_AsyncWait_Wait")));
            //wait until the loading finish
            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.Id("ReportViewer1_AsyncWait_Wait")));
            //retry downloading
            RetryDownload(fullpath);
        }

        public void RetryDownload(string fullpath)
        {
            //retry downloading
            int retryCount = 3;
            while (retryCount > 0)
            {
                //click the save link
                SaveIcon.Click();
                ExcelSaveLink.Click();
                int totalTime = 60000; //60 sec
                bool isFileExists = false;
                //wait for downloading
                while (!(isFileExists = File.Exists(fullpath + ".xlsx")))
                {
                    //if the file does not exitst after downloading
                    //retry it
                    if (totalTime <= 0)
                        break;
                    Thread.Sleep(1000);
                    totalTime -= 1000;
                }
                if (isFileExists)
                    break;
                retryCount--;
            }

        }

    }
}


