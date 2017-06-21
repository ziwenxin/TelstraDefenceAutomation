﻿using System;
using System.Threading;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects.Toll
{
    public class TollLoginPage
    {
        //config sheet
        private ISheet ConfigSheet;

        #region WebElements
        [FindsBy(How = How.Id, Using = "UserName")]
        public IWebElement UserNameField { get; set; }

        [FindsBy(How = How.Id, Using = "Password")]
        public IWebElement PasswordField { get; set; }


        [FindsBy(How = How.XPath, Using = "//button[text()='Login']")]
        public IWebElement LoginBtn { get; set; } 
        #endregion

        /// <summary>
        /// initialize and set config sheet
        /// </summary>
        /// <param name="ConfigSheet"></param>
        public TollLoginPage(ISheet configSheet)
        {
            //initialize
            this.ConfigSheet = configSheet;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);



        }
        /// <summary>
        /// go to the login page of Toll
        /// </summary>
        public void GoToLoginPage()
        {
            IWebDriver driver = WebDriver.ChromeDriver;

            //wait for 2 secs 
            Thread.Sleep(2000);
            //launch the web
            driver.Navigate().GoToUrl(ConfigSheet.GetRow(1).GetCell(1).StringCellValue);

            //wait until the save icon exists
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            wait.Until(ExpectedConditions.ElementExists(By.Id("UserName")));

        }

        /// <summary>
        /// login action
        /// </summary>
        /// <returns></returns>
        public TollReportDownloadPage Login()
        {

            int retryCount = 3;
            //retry 3 times to go to url
            while (true)
            {
                try
                {
                    GoToLoginPage();
                    break;
                }
                catch (Exception e)
                {
                    if (retryCount <= 0)
                        throw e;
                    retryCount--;
                }
            }


            //enter the credentials
            UserNameField.SendKeys(ConfigSheet.GetRow(2).GetCell(1).StringCellValue);
            PasswordField.SendKeys(ConfigSheet.GetRow(3).GetCell(1).StringCellValue);
            //click login
            LoginBtn.Click();
            return new TollReportDownloadPage(ConfigSheet);
        }

    }
}
