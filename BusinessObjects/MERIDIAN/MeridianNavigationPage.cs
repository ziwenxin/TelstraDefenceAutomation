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
        [FindsBy(How = How.XPath, Using = "//a[text()='General Reporting']")]
        public IWebElement GeneralReportLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='Accounts Payable']")]
        public IWebElement AccountPayableLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='Purchasing']")]
        public IWebElement PurchasingLink { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[text()='PO Details']")]
        public IWebElement PODetailsLink { get; set; }

        public MeridianNavigationPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }

        public MeridianVariableEntryPage GotoPoDetailEntryPage(ISheet configSheet)
        {
            //wait general report link exists
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(8));
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
            return new MeridianVariableEntryPage(configSheet);
        }
    }
}
