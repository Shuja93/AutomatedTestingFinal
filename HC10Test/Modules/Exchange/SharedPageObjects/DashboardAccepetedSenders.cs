using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using HC10AutomationFramework.Logs;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;

namespace HC10Test.PageObjects
{
    class DashboardAccepetedSenders : BasePage
    {
        private IWebElement btnAddUsersAcceptedSendersElem => ByXPath("//button[contains(@onclick, 'acceptedsendersUpdate')]");
        public string SetAcceptedSenders(string userList)
        {
            try
            {
                ClickSendersForTheList(DriverContext.Driver, DivContainer.AccpetedSender);
                btnAddUsersAcceptedSendersElem.Click();
                {
                    string addUsers = AddUsersinNewWindows(DriverContext.Driver, userList);
                    if (addUsers != "success")
                    {
                        return addUsers;
                    }
                }
                
                ClickPermissionsSaveButton(DriverContext.Driver, DivContainer.AccpetedSender);
                return GetPrompt( headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


        }
        public string VerifyAcceptedSenders(string userList)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                //WaitforSpinnerbgCondition
                return VerifyUsersInPermissions(DriverContext.Driver, userList, DivContainer.AccpetedSender);

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
