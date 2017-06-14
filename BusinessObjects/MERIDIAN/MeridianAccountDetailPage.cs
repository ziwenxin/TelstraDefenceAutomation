using System;
using System.Collections.Generic;
using System.IO;
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
    public class MeridianAccountDetailPage : MeridianCenterPage
    {
        [FindsBy(How = How.Id, Using = "BUTTON_OPEN_SAVE_btn1_acButton")]
        public IWebElement OpenBtn { get; set; }

        [FindsBy(How = How.Id, Using = "BUTTON_TOOLBAR_2_btn3_acButton")]
        public IWebElement SaveBtn { get; set; }

        [FindsBy(How = How.Id, Using = "FILTER_PANE_ac_feodd_0DOC_DATE_dropdown_combobox")]
        public IWebElement InvoiceDateFilterDpList { get; set; }


        public MeridianAccountDetailPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }

        public MeridianPopUpWindow OpenPoPUpWindow()
        {
            //wait for open button appears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(300));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("BUTTON_OPEN_SAVE_btn1_acButton")));
            //click it
            OpenBtn.Click();

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
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(600));
            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("/html/body/div[13]/img")));


        }


        public MeridianAccDetailFilterWindow AddFilter()
        {
            //wait drop list clickable
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(300));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FILTER_PANE_ac_feodd_0DOC_DATE_dropdown_combobox")));


            //set the filter
            //click the filter, press 'E' then press 'Enter'
            InvoiceDateFilterDpList.Click();
            //wait for 0.5 sec between each key press
            Thread.Sleep(500);
            InvoiceDateFilterDpList.SendKeys("e");
            Thread.Sleep(500);
            InvoiceDateFilterDpList.SendKeys(Keys.Enter);


            //wait for pop up window appears, the id is its body
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            //wait for inner frame         
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("urPopupInner0")));
            WebDriver.ChromeDriver.SwitchTo().Frame(PopUpFrame);
            //wait for the pop up window completed
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("SELECTOR_mainctrl_removeButton")));
            return new MeridianAccDetailFilterWindow();
        }

        public void DownLoadPoDetailDoc(string fullpath)
        {
            //wait generation of the report
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(PODetailInputFrame);
            //wait for loading
            WaitForLoading();


            //dowloading the file
            RetryDownloading(fullpath);

            //change the frame to default
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();

        }


        public void DownLoadAccountDetailDoc(string fullpath)
        {
            //wait generation of the report
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(AccountDetailInputFrame);
            //wait for loading
            WaitForLoading();

            //save the report
            MeridianAccDetailFilterWindow meridianAccDetailFilterWindow = AddFilter();
            meridianAccDetailFilterWindow.AddFilter();
            //switch back 
            //wait generation of the report
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(AccountDetailInputFrame);

            WaitForLoading();


            //download the file
            RetryDownloading(fullpath);

            //change the frame to default
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();

        }

        public void RetryDownloading(string fullpath)
        {
            //retry downloading
            int retryCount = 3;
            while (retryCount > 0)
            {
                //wait for save btn available
                WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(10));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("BUTTON_TOOLBAR_2_btn3_acButton")));
                Thread.Sleep(500);
                //save the report
                SaveBtn.Click();

                //wait downloading of the report
                WaitForLoading();

                int totalTime = 60000; //60 sec
                bool isFileExists = false;
                //wait for downloading
                while (!(isFileExists = File.Exists(fullpath + ".xls")))
                {
                    //if the file does not exitst after downloading
                    //retry it
                    if (totalTime <= 0)
                        break;
                    Thread.Sleep(1000);
                    totalTime -= 1000;
                }
                if (isFileExists)
                    break;
                retryCount--;
            }
        }
    }
}
