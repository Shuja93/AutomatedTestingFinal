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
    class ExgOrgMailboxes : BasePage

    {


        private IWebElement btnCreateMailboxPageElem =>
            ByXPath("//button[contains(@onclick, 'ExgMailbox.AddMailbox')]");

        private IWebElement btnMailboxDashboardElem =>
            ByXPath("//tr[1]//button[contains(@onclick, 'ExgMailbox.ShowDashboard')]");
        private IWebElement searchBarMailboxElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='DisplayName']"));

        private IWebElement btnmailboxDisableElem => ByXPath("//tr[1]//button[contains(@onclick, 'ExgMailbox.EnDisableMailbox(this);')]");

        private IWebElement btnToggle => ByXPath("//*[@id='list-search']//button[contains(@onclick , 'Layout.ShowHideAdvSearch')]");
        private IWebElement txtEmailAddress => ByXPath("//*[@id='EmailAddress']");
        private IWebElement btnAdvanceSearch => ByXPath("//*[@id='AdvSearchSubmitButton']");

        public ExgCreateMailbox OpenCreateMailboxPage()
        {
            btnCreateMailboxPageElem.Click();
            return new ExgCreateMailbox();
        }

        public void SearchMailboxName(string email,string displayName)
        {
            string searchString;

            if (string.IsNullOrEmpty(displayName))
            {
                MailAddress addr = new MailAddress(email);
                searchString = addr.User;
            }
            else
            {
                searchString = displayName;
            }
            

            SeleniumHelperMethods.ObjectSearchBar(DriverContext.Driver, searchBarMailboxElem, btnSearch, headerProgressElem,headerProgressElemBy, searchString);
        }

        public void SearchMailboxUsingEmail(string email) 
        {
            btnToggle.Click();
            txtEmailAddress.SendKeys(email);
            btnAdvanceSearch.Click();
            SeleniumHelperMethods.LoadWait(DriverContext.Driver, headerProgressElem, headerProgressElemBy);
        }

        public ExgMailboxDashboard OpenMailboxDashboard() 
        {
            btnMailboxDashboardElem.ClickWithWait("header");
            return new ExgMailboxDashboard();
        }

        public string DisableMailbox()
        {
            string status;
            try
            {
                
                btnmailboxDisableElem.Click();
                btnVerifyDisableElem.Click();
                //revisit
                status = GetPrompt(  headerProgressElem, headerProgressElemBy, MessageContainer.ToastContainer);
                
            }

            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);
                status = TestStatus.Failed;

            }
            return status;

        }

        public string EnableMailbox()
        {
            string status;
            try
            {

                btnmailboxDisableElem.Click();
                //revisit
                status = GetPrompt( headerProgressElem, headerProgressElemBy, MessageContainer.ToastContainer);
                PageRefresh(DriverContext.Driver);
            }

            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);
                status = TestStatus.Failed;

            }
            return status;

        }



    }
}
