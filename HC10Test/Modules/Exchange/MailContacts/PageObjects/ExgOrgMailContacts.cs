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
     class ExgOrgMailContacts :  BasePage
    {
        private IWebElement btnCreateMailContact =>
            ByXPath("//button[contains(@onclick,'MailContact.AddMailContact')]");
        private IWebElement searchBarMailboxElem =>
            DriverContext.Driver.FindElement(By.XPath("//*[@id='DisplayName']"));

        private IWebElement btnMailContactDashboardElem =>
            ByXPath("//tr[1]//button[contains(@onclick,'MailContact.ShowDashboard')]");

        public ExgCreateMailContact OpenCreateMailContactPage()
        {
            btnCreateMailContact.Click();
            return new ExgCreateMailContact();
        }

        public void SearchMailContact(string email, string displayName)
        {
            string searchString;

            if (displayName == "")
            {
                MailAddress addr = new MailAddress(email);
                searchString = addr.User;
            }
            else
            {
                searchString = displayName;
            }


            SeleniumHelperMethods.ObjectSearchBar(DriverContext.Driver, searchBarMailboxElem, btnSearch, headerProgressElem, headerProgressElemBy, searchString);
        }

        public ExgMailContactDashboard OpenMailContactDashboard()
        {
            btnMailContactDashboardElem.ClickWithWait("header");
            return new ExgMailContactDashboard();
        }




    }
}
