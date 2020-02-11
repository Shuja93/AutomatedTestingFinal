using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using HC10AutomationFramework.Logs;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;

namespace HC10Test.PageObjects
{
    class DashboardFullAccessPermissions : BasePage
    {
        private IWebElement btnAddUsersFullAccessPermissionsElem => ByXPath("//button[contains(@onclick, 'fullaccesspermissioncontainerUpdate')]");
        public string SetFullAccessPermissions(string userList)
        {

            try
            {
                btnAddUsersFullAccessPermissionsElem.Click();
                AddUsersinNewWindows(DriverContext.Driver, userList);
                ClickPermissionsSaveButton(DriverContext.Driver, DivContainer.FullAccessPermissions);
                return GetPrompt( headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }
        public string VerifyFullAccessPermissions(string userList)
        {

            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                //WaitforSpinnerbgCondition
                return VerifyUsersInPermissions(DriverContext.Driver, userList, DivContainer.FullAccessPermissions);

            }
            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);
                return ex.Message;
            }


        }
    }
}
