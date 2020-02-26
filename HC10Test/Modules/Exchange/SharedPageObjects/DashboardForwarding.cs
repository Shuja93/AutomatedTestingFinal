using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using System.Threading;
using HC10AutomationFramework.Logs;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;

namespace HC10Test.PageObjects
{
    class DashboardForwarding : BasePage
    {
        private IWebElement txtForwardingEmail => ByXPath("//*[@id='RecipientName']");
        private IWebElement ckbxEnableForwardingElem => ById("EnableForwardAddress");
        private IWebElement ckbxDeliverAndForwardEnableElem => ById("DeliverAndForwardEnable");
        private IWebElement btnAddUserForwardingElem => ByXPath("//button[contains(@onclick , 'ExgMailboxManager.GetMbxUser')]");
  
        public string SetForwarding(string user, string ou, string exchangeObject, IWebElement forwardingButton)
        {

            try
            {
                SetCheckBox(ckbxEnableForwardingElem,true);
                SetCheckBox(ckbxDeliverAndForwardEnableElem,true);
                
                forwardingButton.Click();
                AddForwardingPopUp(DriverContext.Driver, user, ou, exchangeObject);
                Thread.Sleep(2000);
                ClickPermissionsSaveButton(DriverContext.Driver, DivContainer.Forwarding);
                return GetPrompt( headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


        }
        public string VerifyForwarding(string user)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                Thread.Sleep(2000);
               
                if (DriverContext.Driver.FindElement(By.XPath("//*[@id='RecipientName']")).GetAttribute("value") != user)
                {
                    return TestStatus.Failed;
                }
                else
                {
                    return TestStatus.Success;
                }

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
