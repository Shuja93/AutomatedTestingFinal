using System;
using System.Linq;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Helpers;
using OpenQA.Selenium;

namespace HC10Test.PageObjects
{
     class ExgDLDashboard : BasePage
    {

        private readonly DashboardEmailAddress pageEmailAddress;
        private readonly DashboardSendOnBehalf pageSendOnBehalf;
        private readonly DashboardSendAsPermissions pageSendAsPermissions;
        private readonly DashboardAccepetedSenders pageAcceptedSenders;
        private readonly DashboardRejectedSenders pageRejectedSenders;


        public ExgDLDashboard()
        {
            pageEmailAddress = new DashboardEmailAddress();
            pageSendOnBehalf = new DashboardSendOnBehalf();
            pageSendAsPermissions = new DashboardSendAsPermissions();
            pageAcceptedSenders = new DashboardAccepetedSenders();
            pageRejectedSenders = new DashboardRejectedSenders();

        }

        private IWebElement lnkMembersElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#ListMembers']"));
        private IWebElement lnkAdministratorsElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#Administrator']"));

        private IWebElement btnAddUsersMembersElem =>
            ByXPath("//button[contains(@onclick, 'DistributionList.GetDistributionListMembers')]");
        private IWebElement txtDLDisplayNameElem => ByXPath("//*[@id='DisplayName']");

        private IWebElement btnAddUsersAdministratorsElem =>
            ByXPath("//button[contains(@onclick, 'DistributionList.GetDistributionListAdmins')]");

        private IWebElement isHideFromListElem => ByXPath("//*[@id='IsHideFromList']");
        private IWebElement IsSendOutToOriginatorElem => ById("IsSendOutToOriginator");
        private IWebElement IsAllSenderAuthenticatedElem => ByXPath("//*[@id='IsAllSenderAuthenticated']");
        private IWebElement radioReportToManagerElem => ByXPath("//*[@id='ReportToManager']");
        private IWebElement radioReportToOriginatorElem => ByXPath("//*[@id='ReportToOriginator']");
        private IWebElement radioNoReportElem => ByXPath("//*[@id='NoReport']");
        private IWebElement textAreaNotesElem => ByXPath("//*[@id='Notes']");
        private IWebElement btnAddAdvancePropertie => ByXPath("//*[@id='addMailboxBtn']");
        private IWebElement lnkEmailAddressElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#EmailAddress']"));
        private IWebElement lnkSenAsPermissionsElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SendAs']"));

        private IWebElement ckbxIncomingMessageSizeElem =>
            ByXPath("//*[@id='InComingMessageSizeLimit']//input[@type = 'checkbox']");

        private IWebElement txtIncomingMessageSizeElem =>
            ByXPath("//*[@id='InComingMessageSizeLimit']//input[@type = 'text']");



        public string SetDlMembers(string userList)
        {
            try
            {
                btnAddUsersMembersElem.Click();
                AddMembersInDl(DriverContext.Driver, userList);


                ClickPermissionsSaveButton(DriverContext.Driver, DivContainer.DLMember);
                return GetPrompt(headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string VerifyDlMembers(string members)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                string userList = "";
                var users = members.Split(' ');
                foreach (var user in users)
                {
                    var _user = user.Split(';');
                    userList = userList + _user[1] + ' ';
                }
                
                return VerifyUsersInPermissions(DriverContext.Driver, userList, DivContainer.DLMember);

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string SetDlAdmministrators(string userList)
        {
            try
            {
                btnAddUsersAdministratorsElem.Click();
                AddUsersinNewWindows(DriverContext.Driver, userList);
                ClickPermissionsSaveButton(DriverContext.Driver, DivContainer.DlAdministrator);
                return GetPrompt(headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string VerifyDlAdministrators(string userList)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");

                return VerifyUsersInPermissions(DriverContext.Driver, userList, DivContainer.DlAdministrator);

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string SetAdvanceProperties(string newDisplayName, bool hideFromAddessList,
            bool sendOofMessagetoOriginator, bool sendersInsideandOutsideOrg, string deliveryReport, string notes, string incomingMessageSize,  bool IsCR)
        {
            try
            {
                if (newDisplayName != null)
                {
                    txtDLDisplayNameElem.Clear();
                    txtDLDisplayNameElem.SendKeys(newDisplayName);
                }


                if (hideFromAddessList == true)
                {
                    if (!isHideFromListElem.Selected)
                    {
                        isHideFromListElem.Click();
                    }
                
                }
                else
                {
                    if (isHideFromListElem.Selected)
                    {
                        isHideFromListElem.Click();
                    }
                
                }

                if (sendOofMessagetoOriginator == true)
                {
                    if (!IsSendOutToOriginatorElem.Selected)
                    {
                        IsSendOutToOriginatorElem.Click();
                    }

                }
                else
                {
                    if (IsSendOutToOriginatorElem.Selected)
                    {
                        IsSendOutToOriginatorElem.Click();
                    }

                }

                if (sendersInsideandOutsideOrg == true)
                {
                    if (!IsAllSenderAuthenticatedElem.Selected)
                    {
                        IsAllSenderAuthenticatedElem.Click();
                    }

                }
                else
                {
                    if (IsAllSenderAuthenticatedElem.Selected)
                    {
                        IsAllSenderAuthenticatedElem.Click();
                    }

                }

                switch (deliveryReport)
                {
                    case "Report to Manager":
                        radioReportToManagerElem.Click();
                        break;
                    case "Report to Originator":
                        radioReportToOriginatorElem.Click();
                        break;
                    case "No Delivery Report":
                        radioNoReportElem.Click();
                        break;
                }

                if (notes != null)
                {
                    textAreaNotesElem.Clear();
                    textAreaNotesElem.SendKeys(notes);
                }

                if (incomingMessageSize != null)
                {
                    if (IsCR  == true)
                    {
                        SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, incomingMessageSize);
                    }
                    else
                    {
                        SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, "I'll choose my own offering");
                        if (ckbxIncomingMessageSizeElem.Selected) { ckbxIncomingMessageSizeElem.Click();}
                        txtIncomingMessageSizeElem.Clear();
                        txtIncomingMessageSizeElem.SendKeys(incomingMessageSize);

                    }
                }


                btnAddAdvancePropertie.Click();
                return GetPrompt(headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);

            }
            catch (Exception ex)
            {
                return ex.Message.Trim();
            }




        }



        public string SetAdditionalEmailAddress(string additionalEmail) => pageEmailAddress.SetAdditionalEmailAddress(additionalEmail);
        public string VerifyAdditionalEmailAddress(string additionalEmail) => pageEmailAddress.VerifyAdditionalEmailAddress(additionalEmail);
        public string SetSendOnBehalf(string userList) => pageSendOnBehalf.SetSendOnBehalf(userList);
        public string VerifySendOnBehalf(string userList) => pageSendOnBehalf.VerifySendOnBehalf(userList);

        public string SetSendAsPermissions(string userList) => pageSendAsPermissions.SetSendAsPermissions(userList,DivContainer.SendAsPermissionsDl);
        public string VerifySendAsPermissions(string userList) => pageSendAsPermissions.VerifySendAsPermissions(userList);
        public string SetAcceptedSenders(string userList) => pageAcceptedSenders.SetAcceptedSenders(userList);
        public string VerifyAcceptedSenders(string userList) => pageAcceptedSenders.VerifyAcceptedSenders(userList);
        public string SetRejectedSenders(string userList) => pageRejectedSenders.SetRejectedSenders(userList);
        public string VerifyRejectedSenders(string userList) => pageRejectedSenders.VerifyRejectedSenders(userList);






        public void OpenMembers()
        {
            lnkMembersElem.ClickWithWait("spinner");
        }

        public void OpenAdministrators()
        {
            lnkAdministratorsElem.ClickWithWait("spinner");
        }

        public void OpenEmailAddress()
        {
            lnkEmailAddressElem.ClickWithWait("spinner");
        }
        public void OpenSendAsPermissions()
        {
            lnkSenAsPermissionsElem.ClickWithWait("spinner");
        }

    }
}
