using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects.MERIDIAN
{
    public class MeridianPOAccountDetailPage : MeridianCenterPage
    {
        [FindsBy(How = How.Id, Using = "BUTTON_OPEN_SAVE_btn1_acButton")]
        public IWebElement OpenBtn { get; set; }

        [FindsBy(How = How.Id, Using = "BUTTON_TOOLBAR_2_btn3_acButton")]
        public IWebElement SaveBtn { get; set; }

        [FindsBy(How = How.Id, Using = "FILTER_PANE_ac_feodd_0DOC_DATE_dropdown_combobox-r")]
        public IWebElement InvoiceDateFilterField { get; set; }


        public MeridianPOAccountDetailPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }

        public MeridianPopUpWindow OpenPoPUpWindow()
        {
            //wait for open button appears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(8));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("BUTTON_OPEN_SAVE_btn1_acButton")));
            //click it
            OpenBtn.Click();

            //wait for pop up window appears, the id is its body
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            //wait for inner frame         
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("urPopupInner0")));
            //switch to it 
            WebDriver.ChromeDriver.SwitchTo().Frame(PopUpFrame);
            //wait for the pop up window completed
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("LOAD_state_tigen4_tlv1_list_unid7_tv")));
            return new MeridianPopUpWindow();
        }


        public void WaitForLoading()
        {

            //wait loading image disappears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(120));
            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//img[@src='/com.sap.ip.bi.web.portal.mimes/base/images/generic/pixel.gif?version=AyqckNPrka7NCmWJEfbIYw%3D%3D']")));


        }


        public void AddFilter()
        {
            //wait drop list clickable
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(120));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("LOAD_state_tigen4_tlv1_list_unid7_tv")));

            //get current date
            string nowStr = DateTime.Today.ToString("d").Replace("/", ".");
            //replace the last 3 digit by ...
            nowStr = nowStr.Substring(0, nowStr.Length - 4) + "...";
            //get the date 3 month ago
            string threeMonthAgoStr = DateTime.Today.AddMonths(-3).ToString("d").Replace("/", ".");
            //connect them together, like "1.1.2017 - 1.4.2017";
            string filterStr = threeMonthAgoStr + " - " + nowStr;
            //set the filter
            WebDriver.ChromeDriver.ExecuteJavaScript("document.getElementById('FILTER_PANE_ac_feodd_0DOC_DATE_dropdown_combobox').setAttribute('value','" + filterStr + "')");


        }

        public void DownLoadPoDetailDoc()
        {
            //wait generation of the report
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(PODetailInputFrame);
            //wait for loading
            WaitForLoading();
            //save the report
            SaveBtn.Click();
            //wait downloading of the report
            WaitForLoading();
            //change the frame to default
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();

        }


        public void DownLoadAccountDetailDoc()
        {
            //wait generation of the report
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(AccountDetailInputFrame);
            //wait for loading
            WaitForLoading();
            //save the report
            AddFilter();
            SaveBtn.Click();
            //wait downloading of the report
            WaitForLoading();
            //change the frame to default
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();

        }
    }
}
