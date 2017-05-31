using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace TelstraDefenceAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = ".";
            IWebDriver driver = new ChromeDriver(path);

            driver.Navigate().GoToUrl("https://tcsportal.tollgroup.com/Account/Login");

            IWebElement usernameField = driver.FindElement(By.Id("UserName"));
            IWebElement passwordField = driver.FindElement(By.Id("Password"));
            IWebElement loginBtn = driver.FindElement(By.TagName("button"));


            usernameField.SendKeys("Shahanaz.Syzed@team.telstra.com");
            passwordField.SendKeys("password123");
            loginBtn.Click();



        }
    }
}
