using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using OpenQA.Selenium.Support.UI;
using HC10AutomationFramework.Logs;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Helpers;
using HC10AutomationFramework.Resources;

namespace HC10Test.PageObjects
{
    class DashboardRejectedSenders : BasePage
    {
        private IWebElement btnAddUsersRejectedSendersElem => ByXPath("//button[contains(@onclick, 'rejectedsendersUpdate')]");
        public string SetRejectedSenders(string userList)
        {

            try
            {
                ClickSendersForTheList(DriverContext.Driver, DivContainer.RejectedSender);
                btnAddUsersRejectedSendersElem.Click();
                AddUsersinNewWindows(DriverContext.Driver, userList);
                ClickPermissionsSaveButton(DriverContext.Driver, DivContainer.RejectedSender);
                return GetPrompt( headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (NoSuchElementException)
            {
                return ErrorDescriptions.ErrorAddingUserinPopUp;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


        }
        public string VerifyRejectedSenders(string userList)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                //WaitforSpinnerbgCondition
                return VerifyUsersInPermissions(DriverContext.Driver, userList, DivContainer.RejectedSender);

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
