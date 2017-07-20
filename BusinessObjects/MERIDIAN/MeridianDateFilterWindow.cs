using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects.MERIDIAN
{
    public class MeridianDateFilterWindow : MeridianCenterPage
    {
        #region WebElements

        [FindsBy(How = How.Id, Using = "SELECTOR_mainctrl_componentListControl_unid19_tv")]
        public IWebElement DateRow { get; set; }

        [FindsBy(How = How.Id, Using = "SELECTOR_mainctrl_addButton")]
        public IWebElement AddBtn { get; set; }

        [FindsBy(How = How.Id, Using = "SELECTOR_mainctrl_removeButton")]
        public IWebElement RemoveBtn { get; set; }

        [FindsBy(How = How.Id, Using = "SELECTOR_mainctrl_dropDownForToolSelect_combobox")]
        public IWebElement ValueDpList { get; set; }

        [FindsBy(How = How.Id, Using = "SELECTOR_mainctrl_range_parseInput_inp")]
        public IWebElement InvoiceDateField { get; set; } 
        #endregion

        public MeridianDateFilterWindow()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
        }


        /// <summary>
        /// add date filter, which is 3 month ago to now
        /// </summary>
        public void AddFilter(DateTime startDate)
        {
            //wait for the remove button
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("SELECTOR_mainctrl_componentListControl_unid19_tv")));

            //click the first row
            DateRow.Click();
            Thread.Sleep(500);

            //remove first row
            RemoveBtn.Click();
            Thread.Sleep(500);

            //click show tool
            ValueDpList.Click();

            //wait for 0.5 sec between each key press
            Thread.Sleep(500);
            ValueDpList.SendKeys("v");
            Thread.Sleep(500);
            ValueDpList.SendKeys(Keys.Enter);

            //get current date
            string nowStr = DateTime.Today.ToString("d").Replace("/", ".");
            //get the date 3 month ago
            string threeMonthAgoStr = startDate.ToString("d").Replace("/", ".");
            //connect them together, like "01.01.2017 - 01.04.2017";
            string filterStr = threeMonthAgoStr + " - " + nowStr;
            //wait for invoice date field
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("SELECTOR_mainctrl_range_parseInput_inp")));

            //enter into vendor invoice date
            InvoiceDateField.SendKeys(filterStr);
            //click add button
            AddBtn.Click();
            //click ok
            ClickOkBtn();
        }

    }
}
