using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;

namespace BusinessObjects
{
    public class TollWeb
    {
        //driver for chrome
        private IWebDriver driver;
        //store data from json
        private JObject Jobj;

        [FindsBy(How = How.Id,Using = "UserName")]
        public IWebElement UserNameField { get; set; }

        [FindsBy(How = How.Id, Using = "Password")]
        public IWebElement PasswordField { get; set; }

        [FindsBy(How = How.TagName, Using = "button")]
        public IWebElement LoginBtn { get; set; }

        public TollWeb()
        {
            driver = new ChromeDriver();
            PageFactory.InitElements(driver,this);
            //get json data
            string JSONStr = File.ReadAllText("TollData.json");
            Jobj = JObject.Parse(JSONStr);

        }

        
        public void Login()
        {
            try
            {
                //launch the web
                driver.Navigate().GoToUrl(Jobj["TollURL"].ToString());
                //enter the credentials
                UserNameField.SendKeys(Jobj["UserName"].ToString());
                PasswordField.SendKeys(Jobj["Password"].ToString());
                //click login
                LoginBtn.Click();
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to login");
                
            }
        }

        public void GoToReportPage()
        {
            //launch the web
            driver.Navigate().GoToUrl(Jobj["TollURL"].ToString());
        }
    }
}
