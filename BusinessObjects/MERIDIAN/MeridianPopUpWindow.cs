using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects.MERIDIAN
{
    public class MeridianPopUpWindow :MeridianCenterPage
    {
        #region WebElements
        [FindsBy(How = How.Id, Using = "LOAD_state_tigen4_tlv1_list_unid27_tv")]
        public IWebElement PODetailSpan { get; set; }

        [FindsBy(How = How.Id, Using = "LOAD_state_tigen4_tlv1_list_unid7_tv")]
        public IWebElement AccountDetailSpan { get; set; } 
        #endregion



        public MeridianPopUpWindow()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);

        }

        /// <summary>
        /// select the span of PO Detail and go to the detail page
        /// </summary>
        public void SelectPODetailDoc()
        {
            //wait the span valid
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("LOAD_state_tigen4_tlv1_list_unid27_tv")));
            //select the span
            PODetailSpan.Click();
            //wait for a while
            Thread.Sleep(1000);
            //click OK Button
            clickOkBtn();
        }
        /// <summary>
        /// select the span of Account Detail and go to the detail page
        /// </summary>
        public void SelectAccountDetailDoc()
        {
            //wait the span valid
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(60));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("LOAD_state_tigen4_tlv1_list_unid7_tv")));
            //select the span
            AccountDetailSpan.Click();
            //wait for a while
            Thread.Sleep(1000);
            //click OK Button
            clickOkBtn();
        }


    }
}
