using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using PropertyCollection;

namespace BusinessObjects.SharePoint
{
    public class SharePointPage
    {
        private ISheet ConfigSheet { get; set; }

        [FindsBy(How = How.XPath,Using = "//*[@id='onetidDoclibViewTbl0']/tbody/tr[6]/td[1]/a/img")]
        public IWebElement SaveIcon { get; set; }

        public SharePointPage(ISheet configSheet)
        {
            ConfigSheet = configSheet;
            PageFactory.InitElements(WebDriver.ChromeDriver,this);
        }

        public void DownLoadSharePointDoc()
        {
            //go to the share point page
            RetryNavigation();
           

            //get file full path
            string savePath = ConfigSheet.GetRow(5).GetCell(1).StringCellValue;
            string filename = ConfigSheet.GetRow(34).GetCell(1).StringCellValue;
            filename += ".xlsx";
            //click save 
            RetryDownload(savePath+filename);


        }

        public void RetryNavigation()
        {
            int retryCount = 3;
            //retry 3 times to go to url
            while (true)
            {
                try
                {
                    GoToMainPage();

                    //wait it valid
                    WebDriverWait wait = new WebDriverWait(WebDriver.ChromeDriver, TimeSpan.FromSeconds(30));
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='onetidDoclibViewTbl0']/tbody/tr[6]/td[1]/a/img")));
                    break;
                }
                catch (Exception e)
                {
                    if (retryCount <= 0)
                        throw e;
                    retryCount--;
                }
            }
        }

        private void GoToMainPage()
        {
            //go to share point page
            string URL = ConfigSheet.GetRow(33).GetCell(1).StringCellValue;
            WebDriver.ChromeDriver.Navigate().GoToUrl(URL);
        }

        public void RetryDownload(string fullpath)
        {
            //retry downloading
            int retryCount = 3;
            while (retryCount > 0)
            {
                //click the save link
                SaveIcon.Click();
                int totalTime = 60000; //60 sec
                bool isFileExists = false;
                //wait for downloading
                while (!(isFileExists = File.Exists(fullpath)))
                {
                    //if the file does not exitst after downloading
                    //retry it
                    if (totalTime <= 0)
                        break;
                    Thread.Sleep(1000);
                    totalTime -= 1000;
                }
                if (isFileExists)
                    break;
                retryCount--;
            }

        }
    }
}
