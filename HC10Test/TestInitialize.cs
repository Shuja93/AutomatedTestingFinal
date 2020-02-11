using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Config;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.TestTracker;
using HC10AutomationFramework.Helpers;


namespace HC10Test
{
    public class TestInitialize : BasePage
    {
        public void InitializeSettings()
        {
            if (DriverContext.Driver == null)
            {
                GetDirectory.SetDataDirectory();
                ConfigReader.SetFrameworkSettings();
                ReporterClass report = new ReporterClass();
                report.CreateReporterFile();
                OpenBrowser(Settings.BrowserType);
                DriverContext.Browser.GoToUrl(Settings.AUT);
                DriverContext.Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
                DriverContext.Driver.Manage().Window.Maximize();
                LoginPage login = new LoginPage();
                login.Login();
            }



        }
    }
}
