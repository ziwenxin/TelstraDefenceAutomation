using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects.MERIDIAN
{
    public class MeridianCenterPage
    {
        #region WebElements
        [FindsBy(How = How.Id, Using = "isolatedWorkArea")]
        public IWebElement CenterFrame { get; set; }

        [FindsBy(How = How.Id, Using = "urPopupInner0")]
        public IWebElement PopUpFrame { get; set; }

        [FindsBy(How = How.Id, Using = "iframe_Roundtrip_9223372034830153341")]
        public IWebElement PODetailInputFrame { get; set; }

        [FindsBy(How = How.Id, Using = "iframe_Roundtrip_9223372036154767051")]
        public IWebElement AccountDetailInputFrame { get; set; }

        [FindsBy(How = How.Id, Using = "BUTTON_0")]
        public IWebElement OKBtn { get; set; }

        [FindsBy(How = How.Id, Using = "urPopupOuter0")]
        public IWebElement OutterFrame { get; set; }
        #endregion

        public MeridianCenterPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver,this);
        }
        /// <summary>
        /// click ok button to display the reports in the center main window
        /// </summary>
        public void ClickOkBtn()
        {
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(OutterFrame);
            OKBtn.Click();
        }
    }
}
