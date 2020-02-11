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
    class ExgOrgResourceMailboxes : BasePage

    {


        private IWebElement btnCreateResourceMailboxPageElem =>
            ByXPath("//button[contains(@onclick , 'ResourceMailbox.AddMailbox')]");
        private IWebElement btnResourceMailboxDashboardElem => ByXPath("//tr[1]//button[contains(@onclick , 'ResourceMailbox.ShowDashboard')]");
        private IWebElement searchBarMailboxElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='DisplayName']"));

        protected IWebElement btnmailboxDisableElem => ByXPath("//tr[1]//button[contains(@onclick, 'ExgMailbox.EnDisableMailbox(this);')]");

 
        public ExgCreateResourceMailbox OpenCreateResourceMailboxPage()
        {
            btnCreateResourceMailboxPageElem.Click();
            return  new ExgCreateResourceMailbox();
        }

        public void SearchResourceMailboxName(string email,string displayName)
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
            

            SeleniumHelperMethods.ObjectSearchBar(DriverContext.Driver, searchBarMailboxElem, btnSearch, headerProgressElem,headerProgressElemBy, searchString);
        }

        public ExgResourceMailboxDashboard OpenResourceMailboxDashboard() 
        {
            btnResourceMailboxDashboardElem.ClickWithWait("header");
            return new ExgResourceMailboxDashboard();
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
