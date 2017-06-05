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
            //get configsheet and WebDriver.ChromeDriver
            this.configSheet = Configsheet;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);

            //find elements with the file names from config file
            int totalDocuments = (int) configSheet.GetRow(5).GetCell(1).NumericCellValue;
            if (totalDocuments <= 0)
                throw new NoReportsException();
            //swtich to report frame
            GoToReportPage();
            WebDriver.ChromeDriver.SwitchTo().Frame(ReportFrame);

            //set links
            GoodReportLink = WebDriver.ChromeDriver.FindElement(By.XPath("//a[text()='"+configSheet.GetRow(6).GetCell(1).StringCellValue+"']"));
            ShipDetailLink = WebDriver.ChromeDriver.FindElement(By.XPath("//a[text()='" + configSheet.GetRow(6).GetCell(2).StringCellValue + "']"));
            SOHDetailLink = WebDriver.ChromeDriver.FindElement(By.XPath("//a[text()='" + configSheet.GetRow(6).GetCell(3).StringCellValue + "']"));



        }

        public void GoToReportPage()
        {
            WebDriver.ChromeDriver.Navigate().GoToUrl(configSheet.GetRow(3).GetCell(1).StringCellValue);

        }

        public TollGoodReportPage DownloadGoodDocument()
        {
            //go to report page and click the link

            GoodReportLink.Click();
            return new TollGoodReportPage();
        }

        public TollShipOrderPage DownLoadShipDetail()
        {
            //go to report page and click the link

            ShipDetailLink.Click();
            return new TollShipOrderPage();
        }

        public TollSOHDetailPage DownloadSOHDetail()
        {
            //go to report page and click the link

            SOHDetailLink.Click();
            return new TollSOHDetailPage();
        }
    }

}
