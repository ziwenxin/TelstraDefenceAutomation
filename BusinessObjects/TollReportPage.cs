using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace BusinessObjects
{
    public class TollReportPage
    {
        [FindsBy(How = How.XPath, Using = "(//a[text()='TelDef : Goods Receipt By Date Range'])")]
        public IWebElement GoodReportLink { get; set; }


    }
}
