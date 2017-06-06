﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessObjects;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects
{

    public class TollSOHDetailPage : TollReportPage
    {
        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_txtValue")]
        public IWebElement OwnerIdCbl { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_divDropDown_ctl08")]
        public IWebElement OwnerIdCb { get; set; }


        public TollSOHDetailPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver,this);
        }

        public override void AddFilter()
        {
            //choose owner
            OwnerIdCbl.Click();
            OwnerIdCb.Click();
        }
    }
}