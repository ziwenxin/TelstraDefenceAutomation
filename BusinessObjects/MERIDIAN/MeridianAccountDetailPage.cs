﻿using System;
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

        #region WebElements

        [FindsBy(How = How.Id, Using = "BUTTON_OPEN_SAVE_btn1_acButton")]
        public IWebElement OpenBtn { get; set; }

        [FindsBy(How = How.Id, Using = "BUTTON_TOOLBAR_2_btn3_acButton")]
        public IWebElement SaveBtn { get; set; }

        [FindsBy(How = How.Id, Using = "FILTER_PANE_ac_feodd_0DOC_DATE_dropdown_combobox")]
        public IWebElement InvoiceDateFilterDpList { get; set; } 
        #endregion


        public MeridianAccountDetailPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }
        /// <summary>
        /// click open button to open the document selector
        /// </summary>
        /// <returns>an object of pop up window</returns>
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

        /// <summary>
        /// wait for the loading image disappear
        /// </summary>
        public void WaitForLoading()
        {
            //wait 10 secs for the loading icon
            Thread.Sleep(10000);
            //wait loading image disappears
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(600));
            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("/html/body/div[13]/img")));


        }

        
        /// <summary>
        /// find the date filter and click the 'edit' link
        /// </summary>
        /// <returns>an object of date filter window</returns>
        public MeridianDateFilterWindow AddFilter()
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
            return new MeridianDateFilterWindow();
        }

  

        /// <summary>
        /// download PO Detail Document
        /// </summary>
        /// <param name="fullpath"></param>
        public void DownLoadPoDetailDoc(string fullpath)
        {
            //wait generation of the report
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(PODetailInputFrame);

            //wait for loading
            WaitForLoading();

            //add filter
            MeridianDateFilterWindow meridanDateFilterWindow = AddFilter();
            //The date range for Po Detail should be the 1st of the current year to now
            DateTime startDate = new DateTime(DateTime.Today.Year, 1, 1);
            meridanDateFilterWindow.AddFilter(startDate);
            //wait for a while
            Thread.Sleep(1000);
            //switch back 
            //wait generation of the report
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(PODetailInputFrame);

            WaitForLoading();


            //dowloading the file
            RetryDownloading(fullpath);

            //change the frame to default
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();

        }

        /// <summary>
        /// download account detial document
        /// </summary>
        /// <param name="fullpath"></param>
        public void DownLoadAccountDetailDoc(string fullpath)
        {
            //wait generation of the report
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(AccountDetailInputFrame);
            //wait for loading
            WaitForLoading();

            //Add filter
            MeridianDateFilterWindow meridanDateFilterWindow = AddFilter();
            //the date range for Account Detail should be 3 month ago to current date
            DateTime startDate = DateTime.Today.AddMonths(-3);
            meridanDateFilterWindow.AddFilter(startDate);
            //wait for a while
            Thread.Sleep(1000);
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

        /// <summary>
        /// retry the download action for 3 times 
        /// </summary>
        /// <param name="fullpath"></param>
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
                try
                {

                    SaveBtn.Click();

                }
                catch (Exception e)
                {
                    Console.WriteLine(SaveBtn.Location);
                    Console.WriteLine(SaveBtn.Size);
                    Console.WriteLine(e);
                    SaveBtn.Click();
                    throw e;
                }

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
