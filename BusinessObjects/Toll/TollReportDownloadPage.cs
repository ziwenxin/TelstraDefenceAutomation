using System;
using System.Collections.Generic;
using Exceptions;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects.Toll
{
    public class TollReportDownloadPage
    {
        //config sheet 
        private Dictionary<string,string> ConfigDic;

        #region WebElements

        [FindsBy(How = How.XPath, Using = "//a[text()='TelDef - Goods Receipt By Date Range']")]
        public IWebElement GoodReportLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='TelDef - Shipped Order Report v2']")]
        public IWebElement ShipDetailLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='TelDef - SOH Detail V2']")]
        public IWebElement SOHDetailLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//iframe")]
        public IWebElement ReportFrame { get; set; }
        #endregion

        /// <summary>
        /// initialize and set config sheet
        /// </summary>
        /// <param name="configDic"></param>
        public TollReportDownloadPage(Dictionary<string,string> configDic)
        {
            //get ConfigSheet and WebDriver.ChromeDriver
            this.ConfigDic = configDic;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);

            //find elements with the file names from config file
            int totalDocuments = int.Parse(ConfigDic["TotalTollDocuments"]);
            if (totalDocuments <= 0)
                throw new NoReportsException();
            //switch to report frame
            int retryCount = 3;//retry 3 times if fail to navigate
            while (true)
            {
                try
                {
                    GoToReportPage();
                    break;
                }
                catch (Exception e)
                {
                    if (retryCount <= 0)
                        throw e;
                    retryCount--;
                }

            }



        }
        /// <summary>
        /// go to the report page, which contains all the link of reports
        /// </summary>
        public void GoToReportPage()
        {
            WebDriver.ChromeDriver.Navigate().GoToUrl(ConfigDic["TollReportURL"]);
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(ReportFrame);

        }

        /// <summary>
        /// download good document
        /// </summary>
        /// <returns>an object of good document page</returns>
        public TollGoodReportPage DownloadGoodDocument()
        {
            //wait for the link appears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='" + ConfigDic["TollDocumentName1"] + "']")));
            //click the link
            GoodReportLink.Click();
            return new TollGoodReportPage();
        }
        /// <summary>
        /// download ship order report
        /// </summary>
        /// <returns>an object of ship order page</returns>
        public TollShipOrderPage DownLoadShipOrder()
        {
            //wait for the link appears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='" + ConfigDic["TollDocumentName2"] + "']")));

            //click the link
            ShipDetailLink.Click();
            return new TollShipOrderPage();
        }
        /// <summary>
        /// download SOH detail report
        /// </summary>
        /// <returns>an object of SOH detail page</returns>
        public TollSOHDetailPage DownloadSOHDetail()
        {
            //wait for the link appears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='" + ConfigDic["TollDocumentName3"] + "']")));

            //click on the link
            SOHDetailLink.Click();
            return new TollSOHDetailPage();
        }
    }

}
