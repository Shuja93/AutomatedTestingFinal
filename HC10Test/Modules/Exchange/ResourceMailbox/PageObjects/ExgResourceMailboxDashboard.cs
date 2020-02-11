using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using OpenQA.Selenium.Support.UI;
using HC10AutomationFramework.Logs;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HC10Test.PageObjects
{
    class ExgResourceMailboxDashboard : BasePage
    {
        
        //private IWebElement advancedProperties => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SAdvancedProperties123456']"));
        //private IWebElement btnMailBoxGenPropSaveElem => ById("btnMailBoxGenPropSave");
        //private IWebElement txtGeneralProfileOfficeLocationElem => ById("GeneralProfile_OfficeLocation");
        //private IWebElement txtGeneralProfileStreetAddressElem => ById("GeneralProfile_StreetAddress");
        //private IWebElement txtGeneralProfileCityElem => ById("GeneralProfile_City");
        //private IWebElement txtGeneralProfileZipCodeElem => ById("GeneralProfile_ZipCode");
        //private IWebElement txtGeneralProfileJobTitleElem => ById("GeneralProfile_JobTitle");
        //private IWebElement txtxGeneralProfileCompanyElem => ById("GeneralProfile_Company");
        //private IWebElement btnAddManagerElem => ByXPath("//*[@id='mailboxUserProfileEditor']/div[9]/div[2]/div[2]/button");
        //private IWebElement txtGeneralProfileBusinessPhoneElem => ById("GeneralProfile_BusinessPhone");
        //private IWebElement txtGeneralProfileFaxElem => ById("GeneralProfile_Fax");
        //private IWebElement txtGeneralProfileHomePhoneElem => ById("GeneralProfile_HomePhone");
        //private IWebElement txtGeneralProfilePagerElem => ById("GeneralProfile_Pager");
        //private IWebElement txtGeneralProfileNotesElem => ById("GeneralProfile_Notes");
        //private IWebElement txtGeneralProfileMobilePhoneElem => ById("GeneralProfile_MobilePhone");
        //private IWebElement txtGeneralProfileDepartmentElem => ById("GeneralProfile_Department");
        private IWebElement lnkRetentionPolicyElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SRetentionPolicy123456']"));
        private IWebElement lnkMailboxSignatureElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#MailboxSignature']"));
 
        private IWebElement lnkAutomaticReplyElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#OutOfOffice']"));
        private IWebElement lnkArchiveElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#ArchiveMailbox']"));
        private IWebElement lnkPasswordExpiryElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#MailboxAccountExpiry']"));
        private IWebElement ckbxIsRetentionDefaultElem => ById("IsRetentionDefault");
        private IWebElement txtRetentionDaysElem => ById("RetentionDays");
        private IWebElement btnSaveAddSendOnBehalfUsersElem => ByXPath("//*[@id='sendonbehaldUpdate']/div[3]/div/button");
        private IWebElement listManager => ByXPath("//*[@id='manager']");
        private IWebElement btnAddUserArchiveElem => ByXPath("//button[contains(@onclick, 'ArchiveMailbox')]");
        private IWebElement txtCountryElem => ByXPath("/html/body/span/span/span[1]/input");
        private IWebElement btnSaveRetentionSettingsElem => ByXPath("//*[@id='retentionsettings']//button");
        private IWebElement exchangeObjectEmailList => ByXPath("//*[@id='emailaddresses']//td/span");
        private IWebElement ckbxenableAutoReplyElem => ById("EnableAutoReply");
        private IWebElement ckbxenableExternalMessageElem => ById("SendExternalMessage");
        private IWebElement iFrameMessage => ByXPath("//*[@id='tinymce']/p");
        private IWebElement btnCreateArchiveMailbox => ByXPath("//*[@id='btnSave']");
        
        
        private readonly DashboardEmailAddress pageEmailAddress;
        private readonly DashboardSendOnBehalf pageSendOnBehalf;
        private readonly DashboardFullAccessPermissions pageFullAccessPermissions;
        private readonly DashboardSendAsPermissions pageSendAsPermissions;
        private readonly DashboardAccepetedSenders pageAcceptedSenders;
        private readonly DashboardRejectedSenders pageRejectedSenders;
        private readonly DashboardForwarding pageForwarding;
        private readonly DashboardGeneralProfile pageGeneralProfile;
       private readonly ExgMailboxAdvanceProperties pageAdvanceProperties;

        public ExgResourceMailboxDashboard()
        {
            pageEmailAddress = new DashboardEmailAddress();
            pageSendOnBehalf = new DashboardSendOnBehalf();
            pageFullAccessPermissions = new DashboardFullAccessPermissions();
            pageSendAsPermissions = new DashboardSendAsPermissions();
            pageAcceptedSenders = new DashboardAccepetedSenders();
            pageRejectedSenders = new DashboardRejectedSenders();
            pageForwarding = new DashboardForwarding();
            pageGeneralProfile = new DashboardGeneralProfile();
            pageAdvanceProperties = new ExgMailboxAdvanceProperties();
        }


        

        public string VerifyGeneralProperties(string firstname, string lastName, string displayName, string country, string state, string
                    officeLocation, string address, string city, string zipCode, string jobTitle, string company, string department, string managedBy, string businessPhone, string
                    fax, string homePhone, string mobilePhone, string pager, string notes) => pageGeneralProfile.VerifyGeneralProperties(firstname, lastName, displayName, country, state,
                    officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
                    fax, homePhone, mobilePhone, pager, notes);

        public string VerifyAdvanceProperties(string mailboxSize, bool isCR) => pageAdvanceProperties.VerifyAdvancedProperties(mailboxSize,isCR);




        public string SetGeneralProperties(string firstname, string lastName, string displayName, string country,
            string state, string officeLocation, string address, string city, string zipCode, string jobTitle,
            string company, string department, string managedBy, string businessPhone, string fax, string homePhone,
            string mobilePhone, string pager, string notes) => pageGeneralProfile.SetGeneralProperties(firstname,
            lastName, displayName, country, state,
            officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
            fax, homePhone, mobilePhone, pager, notes);
       
            
            
        
        public string SetRetentionPolicy(string days)
        {
            
            try
            {
                ckbxIsRetentionDefaultElem.Click();
                txtRetentionDaysElem.Clear();
                txtRetentionDaysElem.SendKeys(days);
                btnSaveRetentionSettingsElem.Click();
                return GetPrompt( headerProgressElem, headerProgressElemBy,MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            
        }
        public string VerifyRetentionPolicy(string days)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                //WaitforSpinnerbgCondition
                string daysSet = Convert.ToString(txtRetentionDaysElem.GetAttribute("value"));
                return daysSet == days ? TestStatus.Success : TestStatus.Failed;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            
        }
        public string SetAdditionalEmailAddress(string additionalEmail) => pageEmailAddress.SetAdditionalEmailAddress(additionalEmail);
        public string VerifyAdditionalEmailAddress(string additionalEmail) => pageEmailAddress.VerifyAdditionalEmailAddress(additionalEmail);
        public string SetSendOnBehalf(string userList) => pageSendOnBehalf.SetSendOnBehalf(userList);
        public string VerifySendOnBehalf(string userList) => pageSendOnBehalf.VerifySendOnBehalf(userList);
        public string SetFullAccessPermissions(string userList) => pageFullAccessPermissions.SetFullAccessPermissions(userList);
        public string VerifyFullAccessPermissions(string userList) => pageFullAccessPermissions.VerifyFullAccessPermissions(userList);
        public string SetSendAsPermissions(string userList) => pageSendAsPermissions.SetSendAsPermissions(userList,DivContainer.SendAsPermissions);
        public string VerifySendAsPermissions(string userList) => pageSendAsPermissions.VerifySendAsPermissions(userList);
        public string SetAcceptedSenders(string userList) => pageAcceptedSenders.SetAcceptedSenders(userList);
        public string VerifyAcceptedSenders(string userList) => pageAcceptedSenders.VerifyAcceptedSenders(userList);
        public string SetRejectedSenders(string userList) => pageRejectedSenders.SetRejectedSenders(userList);
        public string VerifyRejectedSenders(string userList) => pageRejectedSenders.VerifyRejectedSenders(userList);
        public string SetForwarding(string user, string ou, string exchangeObject, IWebElement forwardingButton) => pageForwarding.SetForwarding(user, ou, exchangeObject, forwardingButton);
        public string VerifyForwarding(string user) => pageForwarding.VerifyForwarding(user);

        public string SetAutomaticReply (string internalMessage, string externalMessage)
        {

            try
            {
                ckbxenableAutoReplyElem.Click();
                if (externalMessage != null)
                {
                    ckbxenableExternalMessageElem.Click();
                    DriverContext.Driver.SwitchTo().Frame("ExternalMessage_ifr");
                    DriverContext.Driver.FindElement(By.XPath("//*[@id='tinymce']")).SendKeys(externalMessage);
                    DriverContext.Driver.SwitchTo().DefaultContent();
                }
                DriverContext.Driver.SwitchTo().Frame("InternalMessage_ifr");
                DriverContext.Driver.FindElement(By.XPath("//*[@id='tinymce']")).SendKeys(internalMessage);
                DriverContext.Driver.SwitchTo().DefaultContent();
                ClickPermissionsSaveButton(DriverContext.Driver, DivContainer.AutomaticReply);
                return GetPrompt(headerProgressElem, headerProgressElemBy,
                    MessageContainer.ToastContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


        }
        public string VerifyAutomaticReply(string internalMessage, string externalMessage)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");
                //WaitforSpinnerbgCondition

                if (externalMessage != null)
                {
                    ckbxenableExternalMessageElem.Click();
                    DriverContext.Driver.SwitchTo().Frame("ExternalMessage_ifr");
                    if (iFrameMessage.Text != externalMessage)
                    {
                        return TestStatus.Failed;
                    }

                    DriverContext.Driver.SwitchTo().DefaultContent();
                }
                DriverContext.Driver.SwitchTo().Frame("InternalMessage_ifr");
                if (iFrameMessage.Text != internalMessage)
                {
                    return TestStatus.Failed;
                }

                DriverContext.Driver.SwitchTo().DefaultContent();
                return TestStatus.Success;

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string SetArchive()
        {

            try
            {
                btnAddUserArchiveElem.Click();
                SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCreateArchiveMailbox);
                DriverContext.Driver.FindElement(By.XPath("//*[@id='ArchiveQuotaValue']/label/input")).Click();
                btnCreateArchiveMailbox.Click();
                return GetPrompt(headerProgressElem, headerProgressElemBy,
                    MessageContainer.DialogeContainer);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


        }
        public string VerifyArchive(string user)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");

                //WaitforSpinnerbgCondition
                if (DriverContext.Driver.FindElement(By.XPath("//*[@id='ArchiveMbxContainer']//td[1]")).Text
                    .Contains(user))
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


        public void OpenAdvancedProperties()
        {
            lnkadvancedPropertiesElem.ClickWithWait("spinner");
        }

        public void OpenRetentionPolicy()
        {
            lnkRetentionPolicyElem.ClickWithWait("spinner");
        }

        public void OpenMailboxSignature()
        {
            lnkMailboxSignatureElem.ClickWithWait("spinner");
        }

 

        public void OpenMemberships()
        {
            lnkMembershipElem.ClickWithWait("spinner");
        }

        

        public void OpenAutomaticReply()
        {
            lnkAutomaticReplyElem.ClickWithWait("spinner");
        }

        public void OpenArchive()
        {
            lnkArchiveElem.ClickWithWait("spinner");
        }

        public void OpenPasswordExpiry()
        {
            lnkPasswordExpiryElem.Click();
        }
        public void CloseDialogueBox()
        {
            SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCloseDialogueBox);
            Thread.Sleep(2000);
            btnCloseDialogueBox.Click();
        }




        //public IDictionary<string, string> GetStorageQuota()
        //{
        //    IDictionary<string, string> actualMailboxStorageProperties = new Dictionary<string, string>();
        //    try
        //    {
        //        actualMailboxStorageProperties.Add("MailboxSize", Convert.ToString(txtMailboxSizeElem.GetAttribute("value")));
        //        actualMailboxStorageProperties.Add("IncomingSize", Convert.ToString(txtIncomingSize.GetAttribute("value")));
        //        actualMailboxStorageProperties.Add("OutgoingSize", Convert.ToString(txtOutgoingSize.GetAttribute("value")));
        //        actualMailboxStorageProperties.Add("ProhibitSendAt", Convert.ToString(txtProhibitSendAtElem.GetAttribute("value")));
        //        actualMailboxStorageProperties.Add("IssueWarningAt", Convert.ToString(txtWarnAtElem.GetAttribute("value")));
        //        var selectElement = new SelectElement(dropdownCRElem);
        //        string quotaStatus = selectElement.SelectedOption.Text;
        //        actualMailboxStorageProperties.Add("Quota",
        //            quotaStatus == "I'll choose my own offering" ? "Accumulated" : quotaStatus);
        //    }

        //    catch (Exception ex)
        //    {
        //        LogClass.AppendLogs(ex);
        //    }
        //    PageRefresh(DriverContext.Driver);
        //    return actualMailboxStorageProperties;


        //}


        //public IDictionary<string, string> VerifyGeneralProperties()
        //{
        //    try
        //    {
        //        btnTabRefreshButtonElem.Click();
        //        //WaitforSpinnerbgCondition
        //        IDictionary<string, string> actualMailboxGeneralProperties = new Dictionary<string, string>();

        //        actualMailboxGeneralProperties.Add("FirstName", Convert.ToString(txtFirstNameElem.GetAttribute("value")));
        //        actualMailboxGeneralProperties.Add("LastName", Convert.ToString(txtLastNameElem.GetAttribute("value")));
        //        actualMailboxGeneralProperties.Add("DisplayName", Convert.ToString(txtDisplayNameElem.GetAttribute("value")));
        //        actualMailboxGeneralProperties.Add("Country", dropdownCountryElem.Text);
        //        var selectState = new SelectElement(dropdownStateElem);
        //        actualMailboxGeneralProperties.Add("State", selectState.SelectedOption.Text);
        //        actualMailboxGeneralProperties.Add("Office Location", txtGeneralProfileOfficeLocationElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Address", txtGeneralProfileStreetAddressElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("City", txtGeneralProfileCityElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Zip Code", txtGeneralProfileZipCodeElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Job Title", txtGeneralProfileJobTitleElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Company", txtxGeneralProfileCompanyElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Department", txtGeneralProfileDepartmentElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("ManagedBy", listManager.Text);
        //        actualMailboxGeneralProperties.Add("Business Phone", txtGeneralProfileBusinessPhoneElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Fax", txtGeneralProfileFaxElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Home Phone", txtGeneralProfileHomePhoneElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Mobile Phone", txtGeneralProfileMobilePhoneElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Pager", txtGeneralProfilePagerElem.GetAttribute("value"));
        //        actualMailboxGeneralProperties.Add("Notes", txtGeneralProfileNotesElem.GetAttribute("value"));
        //        return actualMailboxGeneralProperties;

        //    }
        //    catch (Exception e)
        //    {
        //        LogClass.AppendLogs(e);
        //        return null;
        //    }
        //}




        //public IDictionary<string, string> GetStorageQuota()
        //{
        //    IDictionary<string, string> actualMailboxStorageProperties = new Dictionary<string, string>();
        //    try
        //    {
        //        actualMailboxStorageProperties.Add("MailboxSize", Convert.ToString(txtMailboxSizeElem.GetAttribute("value")));
        //        actualMailboxStorageProperties.Add("IncomingSize", Convert.ToString(txtIncomingSize.GetAttribute("value")));
        //        actualMailboxStorageProperties.Add("OutgoingSize", Convert.ToString(txtOutgoingSize.GetAttribute("value")));
        //        actualMailboxStorageProperties.Add("ProhibitSendAt", Convert.ToString(txtProhibitSendAtElem.GetAttribute("value")));
        //        actualMailboxStorageProperties.Add("IssueWarningAt", Convert.ToString(txtWarnAtElem.GetAttribute("value")));
        //        var selectElement = new SelectElement(dropdownCRElem);
        //        string quotaStatus = selectElement.SelectedOption.Text;
        //        actualMailboxStorageProperties.Add("Quota",
        //            quotaStatus == "I'll choose my own offering" ? "Accumulated" : quotaStatus);
        //    }

        //    catch (Exception ex)
        //    {
        //        LogClass.AppendLogs(ex);
        //    }
        //    PageRefresh(DriverContext.Driver);
        //    return actualMailboxStorageProperties;


        //}
    }

}
