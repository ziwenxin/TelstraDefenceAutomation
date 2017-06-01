using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects
{
    public class TollLoginPage
    {
        //store data from json
        private ISheet sheet;

        [FindsBy(How = How.Id,Using = "UserName")]
        public IWebElement UserNameField { get; set; }

        [FindsBy(How = How.Id, Using = "Password")]
        public IWebElement PasswordField { get; set; }

        //[FindsBy(How = How.TagName, Using = "button")]

        [FindsBy(How = How.XPath, Using = "//button[text()='Login']")]
        public IWebElement LoginBtn { get; set; }


        public TollLoginPage(ISheet sheet)
        {
            //inital
            this.sheet = sheet;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);



        
        }

        
        public TollReportPage Login()
        {
            try
            {
                IWebDriver driver = WebDriver.ChromeDriver;

                //launch the web
                driver.Navigate().GoToUrl(sheet.GetRow(1).GetCell(0).StringCellValue );
                //enter the credentials
                UserNameField.SendKeys(sheet.GetRow(1).GetCell(1).StringCellValue);
                PasswordField.SendKeys(sheet.GetRow(1).GetCell(2).StringCellValue);
                //click login
                LoginBtn.Click();
                return new TollReportPage(sheet);
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to login");
                Environment.Exit(0);

            }
            return null;
        }

    }
}
