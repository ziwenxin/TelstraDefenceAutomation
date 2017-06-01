using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace PropertyCollection
{
    public static class WebDriver
    {
        public static IWebDriver ChromeDriver { get; set; }

        static WebDriver()
        {
            ChromeDriver = new ChromeDriver();
        }
    }
}
