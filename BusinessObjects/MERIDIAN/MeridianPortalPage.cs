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


    public class MeridianPortalPage
    {
        //config sheet
        private ISheet ConfigSheet { get; set; }


        #region WebElements
        [FindsBy(How = How.Id, Using = "2406890")]
        public IWebElement MeridianLaunchImg { get; set; } 
        #endregion

        public MeridianPortalPage(ISheet configSheet)
        {
            ConfigSheet = configSheet;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);

        }

        /// <summary>
        /// go to the portal and launch meridian
        /// </summary>
        /// <returns>an object of navigation page</returns>
        public MeridianNavigationPage LaunchMeridian()
        {
            int retryCount = 3;
            //retry 3 times
            while (true)
            {
                try
                {
                    //go to launch url
                    WebDriver.ChromeDriver.Navigate().GoToUrl(ConfigSheet.GetRow(10).GetCell(1).StringCellValue);
                    //wait for the image appears
                    WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("2406890")));
                    //click it
                    MeridianLaunchImg.Click();
                    break;
                }
                catch (Exception e)
                {
                    if (retryCount <= 0)
                        throw e;
                    retryCount--;
                }


            }
 
            return new MeridianNavigationPage();
        }
    }
}
