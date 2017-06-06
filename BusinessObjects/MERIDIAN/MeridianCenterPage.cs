using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects.MERIDIAN
{
    public class MeridianCenterPage
    {
        [FindsBy(How = How.Id, Using = "isolatedWorkArea")]
        public IWebElement CenterFrame { get; set; }

        [FindsBy(How = How.Id, Using = "urPopupInner0")]
        public IWebElement PopUpFrame { get; set; }

        [FindsBy(How = How.Id, Using = "iframe_Roundtrip_9223372034830153341")]
        public IWebElement InputFrame { get; set; }

        public MeridianCenterPage()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver,this);
        }
    }
}
