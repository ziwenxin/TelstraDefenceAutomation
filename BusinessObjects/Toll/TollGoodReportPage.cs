using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects
{
    public class TollGoodReportPage :TollReportPage
    {
        #region WebElements

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_txtValue")]
        public IWebElement FromDateField { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl05_txtValue")]
        public IWebElement ToDateField { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_txtValue")]
        public IWebElement OwnerIdCbl { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_divDropDown_ctl08")]
        public IWebElement OwnerIdCb { get; set; } 
        #endregion




        public TollGoodReportPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }

        /// <summary>
        /// add date range, which is from the start of this year to now. Also, click 'TelDef' check box
        /// </summary>
        public override void AddFilter()
        {

            //the data range should be from the first date of this year to today
            int thisYear = DateTime.Now.Year;
            FromDateField.SendKeys(new DateTime(thisYear, 1, 1).ToString());
            ToDateField.SendKeys(DateTime.Today.ToString());
            //check the owner id
            OwnerIdCbl.Click();
            OwnerIdCb.Click();

        }

       




    }
}
