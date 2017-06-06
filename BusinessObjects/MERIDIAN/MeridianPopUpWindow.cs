using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.PageObjects;
using PropertyCollection;

namespace BusinessObjects.MERIDIAN
{
    public class MeridianPopUpWindow :MeridianCenterPage
    {
        public MeridianPopUpWindow()
        {
            PageFactory.InitElements(WebDriver.ChromeDriver, this);

        }
    }
}
