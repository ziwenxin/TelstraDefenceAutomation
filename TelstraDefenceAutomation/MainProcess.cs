using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessObjects;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace TelstraDefenceAutomation
{
    public class MainProcess
    {
        static void Main(string[] args)
        {
            TollWeb tw=new TollWeb();
            tw.Login();

        }



    }
}
