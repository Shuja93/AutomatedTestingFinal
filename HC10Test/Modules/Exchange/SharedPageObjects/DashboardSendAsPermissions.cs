using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using HC10AutomationFramework.Logs;
using HC10AutomationFramework.Enum;
using OpenQA.Selenium.Support.UI;
using HC10AutomationFramework.Extensions;

namespace HC10Test.PageObjects
{
    class DashboardSendAsPermissions : BasePage
    {
        private IWebElement btnAddUsersSendAsPermissionsElem => ByXPath("//button[contains(@onclick, 'sendaspermissionUpdate')]");
        public string SetSendAsPermissions(string userList, string divContainer)
        {

            try
            {
                btnAddUsersSendAsPermissionsElem.Click();
                AddUsersinNewWindows(DriverContext.Driver, userList);
                ClickPermissionsSaveButton(DriverContext.Driver,divContainer);
                return GetPrompt( headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


        }
        public string VerifySendAsPermissions(string userList)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                return VerifyUsersInPermissions(DriverContext.Driver, userList, DivContainer.SendAsPermissions);

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
