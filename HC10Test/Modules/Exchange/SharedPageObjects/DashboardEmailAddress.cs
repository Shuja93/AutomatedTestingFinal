using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using HC10AutomationFramework.Logs;
using System.Net.Mail;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;

namespace HC10Test.PageObjects
{
    class DashboardEmailAddress : BasePage
    {
        private IWebElement btnAddEmailElem => ByXPath("//button[contains(text(),'Add')]");
        private IWebElement txtAddEmailElem => ByXPath("//*[@id='addemailAddress']//*[@id='EmailAddress']");
        private IWebElement btnSaveAddEmailElem => ByXPath("//button[contains(@onclick, 'AddEmailAddress')]");
        public string SetAdditionalEmailAddress(string additionalEmail)
        {
            try
            {
                btnAddEmailElem.Click();
                MailAddress addr = new MailAddress(additionalEmail);
                string userName = addr.User;
                string mailDomain = addr.Host;
                txtAddEmailElem.SendKeys(userName);
                btnSaveAddEmailElem.Click();
                return GetPrompt( headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


        }
        public string VerifyAdditionalEmailAddress(string additionalEmail)
        {
            btnTabRefreshButtonElem.ClickWithWait("spinner");
            //WaitforSpinnerbgCondition
            return CheckEmailAddressPresence(DriverContext.Driver, additionalEmail);
        }

        public string CheckEmailAddressPresence(IWebDriver driver, string emailAddress)
        {
            try
            {
                //revisit-check the total number of email address visible at this path. 
                string findEmailAddress = Convert.ToString(driver.FindElement(
                    By.XPath("//*[@id='emailaddresses']//td/span[contains(text(),'" + emailAddress + "')]")).Text);
                if (findEmailAddress.Trim() == emailAddress)
                {
                    return TestStatus.Success;
                }
                else
                {
                    return TestStatus.Failed;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
