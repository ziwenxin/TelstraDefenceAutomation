﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace BusinessObjects
{
    public abstract class TollReportPage
    {
        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl00")]
        public IWebElement ViewReportBtn { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl05_ctl04_ctl00_ButtonImg")]
        public IWebElement SaveIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Excel']")]
        public IWebElement ExcelSaveLink { get; set; }

        public abstract void AddFilter();
        public abstract void DownLoadReport();

    }
}
