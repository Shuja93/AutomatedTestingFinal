using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using System.Threading;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Helpers;


namespace HC10Test.PageObjects
{
    class ExgResourceMailboxDashboard : BasePage
    {
        private IWebElement txtArchiveMailboxSizeElem => ByXPath("//*[@id='ArchiveQuotaValue']//input[@type='text']");
        private IWebElement txtArchiveWarningLevelElem => ByXPath("//*[@id='ArchiveWarningQuotaValue']");
        private IWebElement ckbxArchiveMailboxSize => ByXPath("//*[@id='ArchiveQuotaValue']//input[@type='checkbox']");
        private IWebElement btnAddForwardingUser => ByXPath("//*[@id='btn-rmbuser']");
        private IWebElement lnkArchiveElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#ArchiveMailbox']"));
        private IWebElement ckbxIsRetentionDefaultElem => ById("IsRetentionDefault");
        private IWebElement txtRetentionDaysElem => ById("RetentionDays");
        private IWebElement btnSaveAddSendOnBehalfUsersElem => ByXPath("//*[@id='sendonbehaldUpdate']/div[3]/div/button");
        private IWebElement listManager => ByXPath("//*[@id='manager']");
        private IWebElement btnAddUserArchiveElem => ByXPath("//button[contains(@onclick, 'ArchiveMailbox')]");
        private IWebElement txtCountryElem => ByXPath("/html/body/span/span/span[1]/input");
        private IWebElement btnSaveRetentionSettingsElem => ByXPath("//*[@id='retentionsettings']//button");
        private IWebElement exchangeObjectEmailList => ByXPath("//*[@id='emailaddresses']//td/span");

        private IWebElement btnCreateArchiveMailbox => ByXPath("//*[@id='btnSave']");
        
        
        private readonly DashboardEmailAddress pageEmailAddress;
        private readonly DashboardSendOnBehalf pageSendOnBehalf;
        private readonly DashboardFullAccessPermissions pageFullAccessPermissions;
        private readonly DashboardSendAsPermissions pageSendAsPermissions;
        private readonly DashboardAccepetedSenders pageAcceptedSenders;
        private readonly DashboardRejectedSenders pageRejectedSenders;
        private readonly DashboardForwarding pageForwarding;
        private readonly DashboardGeneralProfile pageGeneralProfile;
       private readonly ExgResourceMailboxAdvanceProperties pageAdvanceProperties;

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
            pageAdvanceProperties = new ExgResourceMailboxAdvanceProperties();
        }


        

        public string VerifyGeneralProperties(string firstname, string lastName, string displayName, string country, string state, string
                    officeLocation, string address, string city, string zipCode, string jobTitle, string company, string department, string businessPhone, string
                    fax, string homePhone, string mobilePhone, string pager, string notes) => pageGeneralProfile.VerifyGeneralProperties(firstname, lastName, displayName, country, state,
                    officeLocation, address, city, zipCode, jobTitle, company, department, businessPhone,
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
        public string SetForwarding(string user, string ou, string exchangeObject) => pageForwarding.SetForwarding(user, ou, exchangeObject, btnAddForwardingUser);
        public string VerifyForwarding(string user) => pageForwarding.VerifyForwarding(user);

        
        public string SetArchive(string archiveSize, bool isCR)
        {

            try
            {
                btnAddUserArchiveElem.Click();
                SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCreateArchiveMailbox);
                if (isCR == true)
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem,archiveSize);
                }
                else
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, "I'll choose my own offering");
                    SetCheckBox(ckbxArchiveMailboxSize,false);
                    txtArchiveMailboxSizeElem.Clear();
                    txtArchiveMailboxSizeElem.SendKeys(archiveSize);
                    txtArchiveWarningLevelElem.Clear();
                    txtArchiveWarningLevelElem.SendKeys(Convert.ToString(Convert.ToInt64(archiveSize)-1));
                }
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
        public string VerifyArchive(string user,string archiveSize,bool IsCR)
        {
            try
            {
                btnTabRefreshButtonElem.ClickWithWait("spinner");

                //WaitforSpinnerbgCondition
                return DriverContext.Driver.FindElement(By.XPath("//*[@id='ArchiveMbxContainer']//td[1]")).Text
                    .Contains(user) ? TestStatus.Success : TestStatus.Failed;

                DriverContext.Driver.FindElement(By.XPath("//*[@id='ArchiveMbxContainer']//a[contains(@onclick,'ResourceMailbox.EditArchiveMailboxSettings')]")).Click();
                

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

     
        public void OpenMemberships()
        {
            lnkMembershipElem.ClickWithWait("spinner");
        }

        

        public void OpenArchive()
        {
            lnkArchiveElem.ClickWithWait("spinner");
        }


        public void CloseDialogueBox()
        {
            SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCloseDialogueBox);
            Thread.Sleep(2000);
            btnCloseDialogueBox.Click();
        }

    }

}
