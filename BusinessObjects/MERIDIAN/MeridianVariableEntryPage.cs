using System;
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
        public ISheet ConfigSheet { get; set; }

        [FindsBy(How = How.Id,Using = "DLG_VARIABLE_vsc_cvl_VAR_3_INPUT_inp")]
        public IWebElement TelProfitCenterField { get; set; }



        [FindsBy(How = How.Id,Using = "DLG_VARIABLE_dlgBase_BTNOK")]
        public IWebElement OKBtn { get; set; }



        public MeridianVariableEntryPage(ISheet configSheet)
        {
            ConfigSheet = configSheet;
            PageFactory.InitElements(WebDriver.ChromeDriver, this);
            //switch to sub frame
            WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriver.ChromeDriver.SwitchTo().Frame(InputFrame);


        }

        public void WaitForLoading()
        {
            //WebDriver.ChromeDriver.SwitchTo().Frame(CenterFrame);
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(120));
            //wait for it disappears
            //wait for the loading img appears
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("QUERY_TITLE_TextItem")));

        }

        public MeridianPOAccountDetailPage EnterVarible()
        {

            //get code from config file
            string code = ConfigSheet.GetRow(10).GetCell(1).StringCellValue;
            //wait for the input field
            WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(60));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("DLG_VARIABLE_vsc_cvl_VAR_3_INPUT_inp")));
            //input
            TelProfitCenterField.SendKeys(code);
            //wait for ok button
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("DLG_VARIABLE_dlgBase_BTNOK")));
            OKBtn.Click();
            WaitForLoading();
            return new MeridianPOAccountDetailPage();
        }
    }
}
