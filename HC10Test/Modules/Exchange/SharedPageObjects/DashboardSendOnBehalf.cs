using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using HC10AutomationFramework.Logs;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;

namespace HC10Test.PageObjects
{
    class DashboardSendOnBehalf : BasePage
    {
        private IWebElement btnAddUsersSendOnBehalfElem => ByXPath("//button[contains(@onclick, 'sendonbehaldUpdate')]");
        public string SetSendOnBehalf(string userList)
        {

            try
            {
                btnAddUsersSendOnBehalfElem.Click();
                AddUsersinNewWindows(DriverContext.Driver, userList);
                ClickPermissionsSaveButton(DriverContext.Driver, DivContainer.SendOnBehalf);
                return GetPrompt( headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }
        public string VerifySendOnBehalf(string userList)
        {

            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                //WaitforSpinnerbgCondition
                return VerifyUsersInPermissions(DriverContext.Driver, userList, DivContainer.SendOnBehalf);

            }
            catch (Exception ex)
            {
                return ex.Message.Trim();
            }


        }
    }
}
