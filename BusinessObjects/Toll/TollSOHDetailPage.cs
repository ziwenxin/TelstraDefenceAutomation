using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects.Toll
{

    public class TollSOHDetailPage : TollReportPage
    {
        #region WebElement
        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_txtValue")]
        public IWebElement OwnerIdCbl { get; set; }

        [FindsBy(How = How.Id, Using = "ReportViewer1_ctl04_ctl03_divDropDown_ctl08")]
        public IWebElement OwnerIdCb { get; set; } 
        #endregion


        public TollSOHDetailPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver,this);
        }
        /// <summary>
        /// choose the owner
        /// </summary>
        public override void AddFilter()
        {
            //choose owner
            OwnerIdCbl.Click();
            OwnerIdCb.Click();
        }
    }
}
