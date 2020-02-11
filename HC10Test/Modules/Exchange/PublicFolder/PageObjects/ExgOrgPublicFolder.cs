using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Helpers;
using System.Net.Mail;
using HC10AutomationFramework.Enum;
using System;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Logs;


namespace HC10Test.PageObjects
{
     class ExgOrgPublicFolders :  BasePage
     {
        private IWebElement btnCreatePublicFolderElem =>
            ByXPath("//button[contains(@onclick,'PublicFolder.AddPublicFolder')]");
        private IWebElement searchBarMailboxElem =>
            DriverContext.Driver.FindElement(By.XPath("//*[@id='PFName']"));

        private IWebElement btnMailContactDashboardElem =>
            ByXPath("//tr[1]//button[contains(@onclick,'PublicFolder.ShowDashboard')]");

        public ExgCreatePublicFolder OpenCreatePublicFolderPage()
        {
            btnCreatePublicFolderElem.Click();
            return new ExgCreatePublicFolder();
        }

        public void SearchPublicFolder(string email, string displayName)
        {
            string searchString;

            if (string.IsNullOrEmpty(displayName))
            {
                var addr = new MailAddress(email);
                searchString = addr.User;
            }
            else
            {
                searchString = displayName;
            }


            SeleniumHelperMethods.ObjectSearchBar(DriverContext.Driver, searchBarMailboxElem, btnSearch, headerProgressElem, headerProgressElemBy, searchString);
        }

        public ExgPublicFolderDashboard OpenPublicFolderDashboard()
        {
            btnMailContactDashboardElem.ClickWithWait("header");
            return new ExgPublicFolderDashboard();
        }




    }
}
