using System;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Chrome;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Helpers;

using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;

namespace HC10Test
{
    [TestClass]
    public class General : TestBase
    {

        public TestContext TestContext { get; set; }
        public static TestContext _testContext;
        [ClassInitialize]
        public static void ClassSetup(TestContext testContext)
        {
            _testContext = testContext;
        }

        //public TestContext TestContext1 { get; set; }

        
        //string fileName = Environment.CurrentDirectory.ToString() + "\\Data\\Exchange\\ExchangeCreateMailbox.xlsx";

        [TestInitialize]
        public void Setup()
        {
            
            //LogHelpers.CreateLogFile();
            OpenBrowser(BrowserType.Chrome);
            //LogHelpers.Write("Opened Browser:");
            DriverContext.Browser.GoToUrl("https://hostingcontrollerdemo.com/");
            //ExcelHelpers.PopulateInCollection(fileName);
            DriverContext.Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
            DriverContext.Driver.Manage().Window.Maximize();
            

        }

        [TestCleanup]
        public void TearDown()
        {
            DriverContext.Driver.Close();
            DriverContext.Driver.Quit();
        }

       



    }
}
