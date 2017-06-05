using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace BusinessObjects
{
    public class TollShipDetailPage :TollReportPage
    {
        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_txtValue")]
        public IWebElement OrderIDField { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_cbNull")]
        public IWebElement OrderIDCheckBox { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl05_cbNull")]
        public IWebElement FromDateCheckBox { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl07_cbNull")]
        public IWebElement ToDateCheckBox { get; set; }


        public override void AddFilter()
        {
            throw new NotImplementedException();
        }

        public override void DownLoadReport()
        {
            throw new NotImplementedException();
        }
    }
}
