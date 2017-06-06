﻿using System;
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
    public class MeridianPOAccountDetailPage : MeridianCenterPage
    {
        [FindsBy(How = How.Id,Using = "BUTTON_OPEN_SAVE_btn1_acButton")]
        public IWebElement OpenBtn { get; set; }


        public MeridianPOAccountDetailPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver,this);
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
            //WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(PopUpFrame);
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("LOAD_state_tigen4_tlv1_list_unid7_tv")));
            return new MeridianPopUpWindow();
        }

    }
}