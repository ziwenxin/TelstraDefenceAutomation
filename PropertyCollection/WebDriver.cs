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
            
        }

        public static void Init(string path)
        {
            //change the download location
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.default_directory", path);
            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            ChromeDriver = new ChromeDriver(chromeOptions);
        }
    }
}
