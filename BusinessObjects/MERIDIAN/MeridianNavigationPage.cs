using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects.MERIDIAN
{
    public class MeridianNavigationPage
    {
        #region WebElements

        [FindsBy(How = How.XPath, Using = "//a[text()='General Reporting']")]
        public IWebElement GeneralReportLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='Accounts Payable']")]
        public IWebElement AccountPayableLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='Purchasing']")]
        public IWebElement PurchasingLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='Accounting Detail']")]
        public IWebElement AccountDetailLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='PO Details']")]
        public IWebElement PODetailsLink { get; set; }

        #endregion
        public MeridianNavigationPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }

        /// <summary>
        /// Enter the detail and go to PO detail page
        /// </summary>
        /// <returns>an object of variable entry page</returns>
        public MeridianVariableEntryPage GotoPoDetailEntryPage()
        {
            //wait general report link exists
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='General Reporting']")));

            //click on general report link
            GeneralReportLink.Click();
            //wait purchasing link exists
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='Purchasing']")));
            //click on the Purchasing link
            PurchasingLink.Click();
            //wait PO detail link exists
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='PO Details']")));
            PODetailsLink.Click();
            return new MeridianVariableEntryPage();
        }

        /// <summary>
        /// Enter the detail and go to Account detail page
        /// </summary>
        /// <param name="ConfigHelper._configDic"></param>
        /// <returns>an object of variable entry page</returns>
        public MeridianVariableEntryPage GotoAccountDetailEntryPage()
        {
            //wait general report link exists
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(8));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='General Reporting']")));

            //click on general report link
            GeneralReportLink.Click();
            //wait purchasing link exists
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='Accounts Payable']")));
            //click on the Purchasing link
            AccountPayableLink.Click();
            //wait PO detail link exists
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='Accounting Detail']")));
            AccountDetailLink.Click();
            return new MeridianVariableEntryPage();
        }
    }
}
