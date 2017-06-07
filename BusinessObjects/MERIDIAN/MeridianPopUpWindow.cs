using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects.MERIDIAN
{
    public class MeridianPopUpWindow :MeridianCenterPage
    {
        [FindsBy(How = How.Id,Using = "LOAD_state_tigen4_tlv1_list_unid27_tv")]
        public IWebElement PODetailSpan { get; set; }

        [FindsBy(How = How.Id, Using = "LOAD_state_tigen4_tlv1_list_unid7_tv")]
        public IWebElement AccountDetailSpan { get; set; }

        [FindsBy(How = How.Id,Using = "BUTTON_0")]
        public IWebElement OKBtn { get; set; }

        [FindsBy(How = How.Id, Using = "urPopupOuter0")]
        public IWebElement OutterFrame { get; set; }

        public MeridianPopUpWindow()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);

        }

        public void SelectPODetailDoc()
        {
            //wait the span valid
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(8));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("LOAD_state_tigen4_tlv1_list_unid27_tv")));
            //select the span
            PODetailSpan.Click();
            //click OK Button
            clickOkBtn();
        }

        public void SelectAccountDetailDoc()
        {
            //wait the span valid
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(8));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("LOAD_state_tigen4_tlv1_list_unid7_tv")));
            //select the span
            AccountDetailSpan.Click();
            //click OK Button
            clickOkBtn();
        }

        public void clickOkBtn()
        {
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(OutterFrame);
            OKBtn.Click();
        }
    }
}
