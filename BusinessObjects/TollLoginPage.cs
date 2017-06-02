using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects
{
    public class TollLoginPage
    {
        //store data from json
        private ISheet Configsheet;

        [FindsBy(How = How.Id,Using = "UserName")]
        public IWebElement UserNameField { get; set; }

        [FindsBy(How = How.Id, Using = "Password")]
        public IWebElement PasswordField { get; set; }

        //[FindsBy(How = How.TagName, Using = "button")]

        [FindsBy(How = How.XPath, Using = "//button[text()='Login']")]
        public IWebElement LoginBtn { get; set; }


        public TollLoginPage(ISheet Configsheet)
        {
            //inital
            this.Configsheet = Configsheet;
            if (WebDriver.ChromeDriver.Title == "Untitled")
            {
                throw new Exception("Failed to initialize");
            }
            PageFactory.InitElements(WebDriver.ChromeDriver, this);


        
        }

        
        public TollReportDownloadPage Login()
        {

                IWebDriver driver = WebDriver.ChromeDriver;

                //launch the web
                driver.Navigate().GoToUrl(Configsheet.GetRow(0).GetCell(1).StringCellValue);

                //wait until the save icon exists
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                wait.Until(ExpectedConditions.ElementExists(By.Id("UserName")));

                //enter the credentials
                UserNameField.SendKeys(Configsheet.GetRow(1).GetCell(1).StringCellValue);
                PasswordField.SendKeys(Configsheet.GetRow(2).GetCell(1).StringCellValue);
                //click login
                LoginBtn.Click();
                return new TollReportDownloadPage(Configsheet);
        }

    }
}
