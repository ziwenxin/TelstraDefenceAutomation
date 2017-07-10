﻿using System;
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
    public class MeridianVariableEntryPage : MeridianCenterPage
    {
        //config sheet
        private Dictionary<string, string> ConfigDIc { get; set; }
        #region WebElements

        [FindsBy(How = How.Id, Using = "DLG_VARIABLE_vsc_cvl_VAR_3_INPUT_inp")]
        public IWebElement POTelProfitCenterField { get; set; }

        [FindsBy(How = How.Id, Using = "DLG_VARIABLE_vsc_cvl_VAR_2_INPUT_inp")]
        public IWebElement AccountTelProfitCenterField { get; set; }

        [FindsBy(How = How.Id, Using = "DLG_VARIABLE_dlgBase_BTNOK")]
        public IWebElement OKBtn { get; set; } 
        #endregion


        /// <summary>
        /// initial web element and set config sheet
        /// </summary>
        /// <param name="configSheet"></param>
        public MeridianVariableEntryPage(Dictionary<string,string> configDic)
        {
            ConfigDIc = configDic;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);


        }
        /// <summary>
        /// wait for the telstra image which indicates that it has entered the detail page
        /// </summary>
        public void WaitForLoading()
        {
            //WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(300));
            //wait for the telstra img appears
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("QUERY_TITLE_TextItem")));

        }

        /// <summary>
        /// enter variable into the variable page
        /// </summary>
        /// <param name="inputField">the filed element to be inputed in</param>
        /// <param name="inputframe"></param>
        /// <param name="inputId"></param>
        /// <param name="frameId"></param>
        /// <returns>an object of account detial page</returns>
        private MeridianAccountDetailPage EnterVarible(IWebElement inputField,IWebElement inputframe,string inputId,string frameId)
        {
            //switch to certain frame
            SwitchToFrame("isolatedWorkArea", frameId, inputframe);

            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(600));


            //get code from config file
            string code = ConfigDIc["TelstraProfitCentres"];
            //wait for the input field
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id(inputId)));
            //input
            inputField.SendKeys(code);
            //wait for ok button
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("DLG_VARIABLE_dlgBase_BTNOK")));
            OKBtn.Click();
            WaitForLoading();
            return new MeridianAccountDetailPage();
        }

        private void SwitchToFrame(string frameId1,string frameId2,IWebElement frame2)
        {
            //switch to correct frame
            WebDriver.ChromeDriver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(120));

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id(frameId1)));
            //switch to sub frame
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            //wait centre frame
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id(frameId2)));
            WebDriver.ChromeDriver.SwitchTo().Frame(frame2);
        }
         /// <summary>
        /// the entry method of account detail
        /// </summary>
        /// <returns>account detail page</returns>
        public MeridianAccountDetailPage AccountEnter()
        {
            return EnterVarible(AccountTelProfitCenterField, AccountDetailInputFrame, "DLG_VARIABLE_vsc_cvl_VAR_2_INPUT_inp",
                "iframe_Roundtrip_9223372036154767051");
        }
        /// <summary>
        /// the entry method for po detail
        /// </summary>
        /// <returns>Po detail page</returns>
        public MeridianAccountDetailPage PODetailEnter()
        {
            return EnterVarible(POTelProfitCenterField, PODetailInputFrame, "DLG_VARIABLE_vsc_cvl_VAR_3_INPUT_inp",
                "iframe_Roundtrip_9223372034830153341");

        }
    }
}
