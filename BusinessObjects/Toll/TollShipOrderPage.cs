using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects
{
    public class TollShipOrderPage :TollReportPage
    {

        #region WebElements

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl05_txtValue")]
        public IWebElement FromDateField { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_txtValue")]
        public IWebElement ToDateField { get; set; }


        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl05_cbNull")]
        public IWebElement FromDateCheckBox { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_cbNull")]
        public IWebElement ToDateCheckBox { get; set; } 
        #endregion

        public TollShipOrderPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);

        }

        /// <summary>
        /// add date from the start of year to now
        /// </summary>
        public override void AddFilter()
        {
            //uncheck the null boxes
            FromDateCheckBox.Click();
            ToDateCheckBox.Click();
            //the date range
            int thisYear = DateTime.Now.Year;
            FromDateField.SendKeys(new DateTime(thisYear, 1, 1).ToString());
            ToDateField.SendKeys(DateTime.Today.ToString());
        }


    }
}


