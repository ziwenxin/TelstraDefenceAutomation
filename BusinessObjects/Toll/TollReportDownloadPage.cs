﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Exceptions;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;
using OpenQA.Selenium.Support.UI;

namespace BusinessObjects
{
    public class TollReportDownloadPage
    {
        private ISheet configSheet;


        [FindsBy(How = How.XPath, Using = "//a[text()='TelDef - Goods Receipt By Date Range']")]
        public IWebElement GoodReportLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='TelDef - Shipped Order Report v2']")]
        public IWebElement ShipDetailLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='TelDef - SOH Detail V2']")]
        public IWebElement SOHDetailLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//iframe")]
        public IWebElement ReportFrame { get; set; }

        public TollReportDownloadPage(ISheet Configsheet)
        {
            //get configsheet and WebDriver.ChromeDriver
            this.configSheet = Configsheet;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);

            //find elements with the file names from config file
            int totalDocuments = (int)configSheet.GetRow(5).GetCell(1).NumericCellValue;
            if (totalDocuments <= 0)
                throw new NoReportsException();
            //swtich to report frame
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

        public void GoToReportPage()
        {
            WebDriver.ChromeDriver.Navigate().GoToUrl(configSheet.GetRow(3).GetCell(1).StringCellValue);
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(ReportFrame);

        }

        public TollGoodReportPage DownloadGoodDocument()
        {
            //wait for the link appears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='" + configSheet.GetRow(6).GetCell(1).StringCellValue + "']")));
            //click the link
            GoodReportLink.Click();
            return new TollGoodReportPage();
        }

        public TollShipOrderPage DownLoadShipOrder()
        {
            //wait for the link appears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='" + configSheet.GetRow(6).GetCell(2).StringCellValue + "']")));

            //click the link
            ShipDetailLink.Click();
            return new TollShipOrderPage();
        }

        public TollSOHDetailPage DownloadSOHDetail()
        {
            //wait for the link appears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='" + configSheet.GetRow(6).GetCell(3).StringCellValue + "']")));

            //click on the link
            SOHDetailLink.Click();
            return new TollSOHDetailPage();
        }
    }

}
