using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;

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
            //enbale pop up windows
            chromeOptions.AddArgument("test-type");
            chromeOptions.AddArgument("disable-popup-blocking");
            DesiredCapabilities capabilities = DesiredCapabilities.Chrome();
            capabilities.SetCapability(ChromeOptions.Capability, chromeOptions);
            //set default timeout to 5 minutes
            ChromeDriver = new ChromeDriver(".",chromeOptions,new TimeSpan(0,5,0));

        }
    }
}
