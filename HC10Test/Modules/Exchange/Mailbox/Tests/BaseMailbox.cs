using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Logs;
using HC10AutomationFramework.Enum;
using System.Net.Mail;
using System.Threading;
using HC10AutomationFramework.Config;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Helpers;
using HC10AutomationFramework.TestTracker;
using OpenQA.Selenium;

namespace HC10Test
{
    public class BaseMailbox : BasePage

    {
        private readonly ExgMailboxDashboard pageMailboxDashboard;

        public BaseMailbox()
        {
            pageMailboxDashboard = new ExgMailboxDashboard();
        }
        public string CreateMailbox(TestContext testContext)
        {
            try
            {
                //revisit - VerifyOUMethod
                //Stage
                string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
                string mailboxType = Convert.ToString(testContext.DataRow["MailboxType"]);
                bool isSubOU = Convert.ToBoolean(testContext.DataRow["IsSubOU"]);
                string email = Convert.ToString(testContext.DataRow["Email"]);
                bool isNewUser = Convert.ToBoolean(testContext.DataRow["IsNewUser"]);
                string mailboxPassword = Convert.ToString(testContext.DataRow["MailboxPassword"]);
                string mailboxSize = Convert.ToString(testContext.DataRow["MailboxSize"]);
                bool isCR = Convert.ToBoolean(testContext.DataRow["IsCR"]);
                bool passwordChange = Convert.ToBoolean(testContext.DataRow["IsChangePassword"]);
                string firstname = Convert.ToString(testContext.DataRow["FirstName"]);
                string lastName = Convert.ToString(testContext.DataRow["LastName"]);
                string displayName = Convert.ToString(testContext.DataRow["DisplayName"]);
                string country = Convert.ToString(testContext.DataRow["Country"]);
                string state = Convert.ToString(testContext.DataRow["State"]);
                string officeLocation = Convert.ToString(testContext.DataRow["OfficeLocation"]);
                string address = Convert.ToString(testContext.DataRow["Address"]);
                string city = Convert.ToString(testContext.DataRow["City"]);
                string zipCode = Convert.ToString(testContext.DataRow["ZipCode"]);
                string jobTitle = Convert.ToString(testContext.DataRow["JobTitle"]);
                string company = Convert.ToString(testContext.DataRow["Company"]);
                string department = Convert.ToString(testContext.DataRow["Department"]);
                string managedBy = Convert.ToString(testContext.DataRow["ManagedBy"]);
                string businessPhone = Convert.ToString(testContext.DataRow["BusinessPhone"]);
                string fax = Convert.ToString(testContext.DataRow["Fax"]);
                string homePhone = Convert.ToString(testContext.DataRow["HomePhone"]);
                string mobilePhone = Convert.ToString(testContext.DataRow["MobilePhone"]);
                string pager = Convert.ToString(testContext.DataRow["Pager"]);
                string notes = Convert.ToString(testContext.DataRow["Notes"]);


                //Act
                ExgOrgMailboxes pageOrgMailboxes = new ExgOrgMailboxes();
                ExgCreateMailbox pageCreateMailbox = pageOrgMailboxes.OpenCreateMailboxPage();
                string standing = pageCreateMailbox.CreateMailbox(mailboxType, isSubOU, isNewUser, email, mailboxPassword, isCR,
                    mailboxSize, passwordChange, firstname, lastName, displayName, country, state, officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
                    fax, homePhone, mobilePhone, pager, notes);


                //Verify
                var status = VerifyResult(ExchangeMessages.CreateMailbox, standing);
                if (status == TestStatus.Failed)
                {
                   CloseDialogueBox();
                }
                else
                {
                    Thread.Sleep(5000);
                }
                ReporterClass.Reporter("Exchange", "Host", "Create Mailbox", "Mailbox Creation Test", organizationName, "Mailbox", email, "SubOU: " + isSubOU + "; IsNewUser: " + isNewUser + "; IsCr: " + isCR + "; Mailbox/CR Size :" + mailboxSize, status, standing);
                TestTracker.mailboxStatus.Add(email, status);

                return status;
            }
            catch (Exception e)
            {
                LogClass.AppendLogs(e.Message);
                return TestStatus.Failed;
            }
            
        }



        public string VerifyMailBoxGeneralProfile(TestContext testContext,bool isNewMailbox)
        {
            //Stage
            
            string status = TestStatus.Success;
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string firstname = Convert.ToString(testContext.DataRow["FirstName"]);
            string lastName = Convert.ToString(testContext.DataRow["LastName"]);
            string displayName;
            if (isNewMailbox) {  displayName = Convert.ToString(testContext.DataRow["DisplayName"]); }
            else
            {
                displayName = Convert.ToString(testContext.DataRow["NewDisplayName"]);
            }
            
            string country = Convert.ToString(testContext.DataRow["Country"]);
            string state = Convert.ToString(testContext.DataRow["State"]);
            string officeLocation = Convert.ToString(testContext.DataRow["OfficeLocation"]);
            string address = Convert.ToString(testContext.DataRow["Address"]);
            string city = Convert.ToString(testContext.DataRow["City"]);
            string zipCode = Convert.ToString(testContext.DataRow["ZipCode"]);
            string jobTitle = Convert.ToString(testContext.DataRow["JobTitle"]);
            string company = Convert.ToString(testContext.DataRow["Company"]);
            string department = Convert.ToString(testContext.DataRow["Department"]);
            string managedBy = Convert.ToString(testContext.DataRow["ManagedBy"]);
            string businessPhone = Convert.ToString(testContext.DataRow["BusinessPhone"]);
            string fax = Convert.ToString(testContext.DataRow["Fax"]);
            string homePhone = Convert.ToString(testContext.DataRow["HomePhone"]);
            string mobilePhone = Convert.ToString(testContext.DataRow["MobilePhone"]);
            string pager = Convert.ToString(testContext.DataRow["Pager"]);
            string notes = Convert.ToString(testContext.DataRow["Notes"]);

            
            
            //Act
            string standing = pageMailboxDashboard.VerifyGeneralProperties(firstname, lastName, displayName, country, state,
                officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
                fax, homePhone, mobilePhone, pager, notes);

            //Verify
            if (!string.IsNullOrEmpty(standing)) 
            { 
                status = TestStatus.Failed;
            }

            if (isNewMailbox)
            {
                ReporterClass.Reporter("Exchange", "Host", "Verify New mailbox General Properties", "Test to verify that the General Properties set at the time of mailbox creation are set successfully", organizationName, "Mailbox", email, "", status, standing);
            }
            else
            {
                ReporterClass.Reporter("Exchange", "Host", "Verify New mailbox General Properties", "Test to verify that the General Properties set at the time of mailbox update are set successfully", organizationName, "Mailbox", email, "", status, standing);
            }
            return status;
        }

        public string VerifyMailBoxAdvanceProperties(TestContext testContext, bool isNewMailbox)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string mailboxSize = Convert.ToString(testContext.DataRow["MailboxSize"]);
            bool isCR = Convert.ToBoolean(testContext.DataRow["IsCR"]);

            //Act

            //if (isNewMailbox)
            //{
            //    pageMailboxDashboard.OpenAdvancedProperties();
                
            //}

            pageMailboxDashboard.OpenAdvancedProperties();
            btnTabRefreshButtonElem.ClickWithWait("spinner");
            
            string standing = pageMailboxDashboard.VerifyAdvanceProperties(mailboxSize,isCR);
        

    
            //Verify
            if (!string.IsNullOrWhiteSpace(standing))
            {
                status = TestStatus.Failed;
            }

            if (isNewMailbox)
            {
                ReporterClass.Reporter("Exchange", Settings.UserLevel, "Verify New mailbox Advance Properties", "Test to verify that the Advance Properties set at the time of mailbox creation are set successfully", organizationName, "Mailbox", email, "", status, standing);

            }
            else
            {
                ReporterClass.Reporter("Exchange", Settings.UserLevel, "Verify New mailbox Advance Properties", "Test to verify that the Advance Properties set at the time of mailbox update are set successfully", organizationName, "Mailbox", email, "", status, standing);
            }
            return status;
        }

        public string DisableMailbox(TestContext testContext)
        {

            try
            {
                //Arrange
                string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
                string email = Convert.ToString(testContext.DataRow["Email"]);
                string displayName = Convert.ToString(testContext.DataRow["DisplayName"]);
                //Act
                HomePage home = new HomePage();
                home.ClickProvisioning();
                ExchangeHome exgHome = home.ClickExchangeHome();
                exgHome.SearchOrganizationName(organizationName);
                ExgOrgMailboxes orgMailboxes = exgHome.MailboxesHome();
                orgMailboxes.SearchMailboxName(email, displayName);
                return orgMailboxes.DisableMailbox();
            }

            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);
                return null;
            }
        }

        public string EnableMailbox(TestContext testContext)
        {

            try
            {

                //Arrange
                string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
                string email = Convert.ToString(testContext.DataRow["Email"]);
                string displayName = Convert.ToString(testContext.DataRow["DisplayName"]);
                //Act
                HomePage home = new HomePage();
                home.ClickProvisioning();
                ExchangeHome exgHome = home.ClickExchangeHome();
                exgHome.SearchOrganizationName(organizationName);
                ExgOrgMailboxes orgMailboxes = exgHome.MailboxesHome();
                orgMailboxes.SearchMailboxName(email, displayName);
                return orgMailboxes.DisableMailbox();
            }

            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);
                return null;
            }
        }

        public string UpdateMailboxGeneralProperties(TestContext testContext)
        {
            
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string firstname = Convert.ToString(testContext.DataRow["FirstName"]);
            string lastName = Convert.ToString(testContext.DataRow["LastName"]);
            string displayName = Convert.ToString(testContext.DataRow["NewDisplayName"]);
            string country = Convert.ToString(testContext.DataRow["Country"]);
            string state = Convert.ToString(testContext.DataRow["State"]);
            string officeLocation = Convert.ToString(testContext.DataRow["OfficeLocation"]);
            string address = Convert.ToString(testContext.DataRow["Address"]);
            string city = Convert.ToString(testContext.DataRow["City"]);
            string zipCode = Convert.ToString(testContext.DataRow["ZipCode"]);
            string jobTitle = Convert.ToString(testContext.DataRow["JobTitle"]);
            string company = Convert.ToString(testContext.DataRow["Company"]);
            string department = Convert.ToString(testContext.DataRow["Department"]);
            string managedBy = Convert.ToString(testContext.DataRow["ManagedBy"]);
            string businessPhone = Convert.ToString(testContext.DataRow["BusinessPhone"]);
            string fax = Convert.ToString(testContext.DataRow["Fax"]);
            string homePhone = Convert.ToString(testContext.DataRow["HomePhone"]);
            string mobilePhone = Convert.ToString(testContext.DataRow["MobilePhone"]);
            string pager = Convert.ToString(testContext.DataRow["Pager"]);
            string notes = Convert.ToString(testContext.DataRow["Notes"]);
            
            
            //Act
            string standing = pageMailboxDashboard.SetGeneralProperties(firstname, lastName, displayName,country, state,
                officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
                fax, homePhone, mobilePhone, pager, notes);
        

      
            //Verify
            string status = VerifyResult(ExchangeMessages.UpdateMailboxGeneralProperties, standing);
            ReporterClass.Reporter("Exchange", "Host", "Update Mailbox General Properties", "Test to verify if General Properties are being updated properly or not", organizationName, "Mailbox", email, "Refer to CSV File", status, standing);
            return status;
        }

        public string UpdateMailboxAdvanceProperties(TestContext testContext)
        {

            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            bool isCR = Convert.ToBoolean(testContext.DataRow["IsCR"]);
            string mailboxSize = Convert.ToString(testContext.DataRow["MailboxSize"]);
            bool isHiddenFromAddressBook = Convert.ToBoolean(testContext.DataRow["IsHiddenFromAddressBook"]);
            bool isImapEnabled = Convert.ToBoolean(testContext.DataRow["IMAP"]);
            bool isPopEnabled = Convert.ToBoolean(testContext.DataRow["POP"]);
            bool isOwaEnabled = Convert.ToBoolean(testContext.DataRow["OWA"]);
            bool isMapiEnabled = Convert.ToBoolean(testContext.DataRow["MAPI"]);

            pageMailboxDashboard.OpenAdvancedProperties();

            //Act
            string standing = pageMailboxDashboard.SetAdvanceProperties(isCR, mailboxSize,  isHiddenFromAddressBook,  isImapEnabled,  isPopEnabled,  isOwaEnabled,  isMapiEnabled);



            //Verify
            string status = VerifyResult(ExchangeMessages.UpdateUserMailboxAdvanceProperties, standing);
            ReporterClass.Reporter("Exchange", "Host", "Update Mailbox General Properties", "Test to verify if General Properties are being updated properly or not", organizationName, "Mailbox", email, "Refer to CSV File", status, standing);
            return status;
        }

        public string UpdateRetentionPolicy(TestContext testContext)
        {
            
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string retentionPolicyDays = Convert.ToString(testContext.DataRow["RetentionPolicyDays"]);
            
            //Act
            pageMailboxDashboard.OpenRetentionPolicy();
            string standing = pageMailboxDashboard.SetRetentionPolicy(retentionPolicyDays);
        
        
        
            //Verify
            string status = VerifyResult(ExchangeMessages.UpdateRetentionPolicy, standing);
            ReporterClass.Reporter("Exchange", "Host", "Update Mailbox Retention Policy", "Test to check if Retention Policy is being set properly or not", organizationName, "Mailbox", email, "Retain Deleted Items For: " + retentionPolicyDays, status, standing);
            return status;
        }

        public string VerifyRetentionPolicy(TestContext testContext)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string RetentionPolicyDays = Convert.ToString(testContext.DataRow["RetentionPolicyDays"]);

            //Act
            string standing = pageMailboxDashboard.VerifyRetentionPolicy(RetentionPolicyDays);
        
       
            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Update Mailbox Retention Policy", "Test to check if the Retention Policy has been set properly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;

        }

        public string AddAdditionalEmailAddress(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string newEmail = Convert.ToString(testContext.DataRow["AdditionalEmailAddress"]);
           
            //Act
            pageMailboxDashboard.OpenEmailAddress();
            string standing = pageMailboxDashboard.SetAdditionalEmailAddress(newEmail);
        
       
        
            //Verify
            string status = VerifyResult(ExchangeMessages.AddEmailAddress, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Email Address", "Test to check if email addresses are added as additional aliases or not", organizationName, "Mailbox", email, "Email: " + newEmail, status, standing);
            
            return status;

        }

        public string VerifyAdditionalEmailAddress(TestContext testContext)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string newEmail = Convert.ToString(testContext.DataRow["AdditionalEmailAddress"]);
            
            //Act
            string standing = pageMailboxDashboard.VerifyAdditionalEmailAddress(newEmail);
        
       
            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of new email alias", "Test to check if the email address has been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;
        }

        public string AddSendOnBehalfUsers(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["SendOnBehalfUsers"]);

            //Act
            pageMailboxDashboard.OpenSendOnBehalf();
            string standing = pageMailboxDashboard.SetSendOnBehalf(userList);

            //Verify
            string status = VerifyResult(ExchangeMessages.AddSendOnBehalfUsers, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Send On Behalf Users", "Test to check if Send On Behalf users are added successfully", organizationName, "Mailbox", email, "Email List: " + userList, status, standing);
        
            return status;
        }

        public string VerifyAddSendOnBehalfUsers(TestContext testContext)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["SendOnBehalfUsers"]);

            //Act
            string standing = pageMailboxDashboard.VerifySendOnBehalf(userList);
       
            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of new Send On Behalf users", "Test to check if the Send On Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
        
            return status;
        }

        public string AddFullAccessPermissions(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["SendOnBehalfUsers"]);
            
            //Act
            pageMailboxDashboard.OpenFullAccessPermissions();
            string standing = pageMailboxDashboard.SetFullAccessPermissions(userList);
        

        
            //Verify
            string status = VerifyResult(ExchangeMessages.AddFullAccessPermissions, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Full Access Permissions", "Test to check if Full Access Permissions are being added successfully", organizationName, "Mailbox", email, "Email List: " + userList, status, standing);
            
            return status;
        }

        public string VerifyFullAccessPermissions(TestContext testContext)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["FullAccessPermissionsUsers"]);
            
            //Act
            string standing = pageMailboxDashboard.VerifyFullAccessPermissions(userList);
      
            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Full Access Permissions", "Test to check if Full Access Permissions have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            
            return status;
        }

        public string AddSendAsPermissions(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["SendAsPermissionsUsers"]);
            pageMailboxDashboard.OpenSendAsPermissions();
            
            //Act
            string standing = pageMailboxDashboard.SetSendAsPermissions(userList);

            //Verify
            string status = VerifyResult(ExchangeMessages.AddSendAsPermissions, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Send As Permissions", "Test to check if Send As Permissions are being added successfully", organizationName, "Mailbox", email, "Email List: " + userList, status, standing);
            return status;


        }

        public string VerifySendAsPermissions(TestContext testContext)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["SendAsPermissionsUsers"]);

            //Act
            string standing = pageMailboxDashboard.VerifySendAsPermissions(userList);


            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Send As Permissions", "Test to check if Send As Permissions have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;
        }

        public string AddAcceptedSenders(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["AcceptedSenders"]);
            pageMailboxDashboard.OpenAcceptedSenders();

            //Act
            string standing = pageMailboxDashboard.SetAcceptedSenders(userList);
            
            //Verify
            string status = VerifyResult(ExchangeMessages.AddAcceptedUsers, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Accepted Users", "Test to check if Accepted Users are being added successfully", organizationName, "Mailbox", email, "Email List: " + userList, status, standing);
            return status;

        }

        public string VerifyAcceptedSenders(TestContext testContext)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["AcceptedSenders"]);
               
            //Act
            string standing = pageMailboxDashboard.VerifyAcceptedSenders(userList);

            //Verify
            if(standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Accepted Users", "Test to check if Accepted Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;

        }

        public string AddRejectedSenders(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["RejectedSenders"]);
            pageMailboxDashboard.OpenRejectedSenders();

            //Act
            string standing = pageMailboxDashboard.SetRejectedSenders(userList);

            //Verify    
            string status = VerifyResult(ExchangeMessages.AddRejectedSenders, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Rejected Users", "Test to check if Rejected Users are being added successfully", organizationName, "Mailbox", email, "Email List: " + userList, status, standing);
            return status;

        }

        public string VerifyRejectedSenders(TestContext testContext)
        {
            string status = TestStatus.Success;
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["RejectedSenders"]);

                string standing= pageMailboxDashboard.VerifyRejectedSenders(userList);

                if (standing != TestStatus.Success)
                {
                    status = TestStatus.Failed;
                }
                ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Rejected Users", "Test to check if Rejected Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
                return status;


        }

        public string AddForwarding(TestContext testContext)
        {
            
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string ou = Convert.ToString(testContext.DataRow["ForwardingOU"]);
            string exchangeObject = Convert.ToString(testContext.DataRow["ForwardingObject"]);
            string user = Convert.ToString(testContext.DataRow["ForwardingEmail"]);

            pageMailboxDashboard.OpenForwarding();
            string standing = pageMailboxDashboard.SetForwarding(user,ou,exchangeObject);

            string status = VerifyResult(ExchangeMessages.AddForwarding, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Forwarding", "Test to check if Forwarding User is being added successfully", organizationName, "Mailbox", email, "Organization: " + ou+ "; Exchange Object: "+exchangeObject+ "; Email: "+user,status, standing);
            return status;

        }

        public string VerifyForwarding(TestContext testContext)
        {
            string status = TestStatus.Success;
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string user = Convert.ToString(testContext.DataRow["ForwardingEmail"]);

            string standing =  pageMailboxDashboard.VerifyForwarding(user);

            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Forwarding User", "Test to check if forwarding Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;
        }

        public string AddAutomaticReply(TestContext testContext)
        {
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string internalMessage = Convert.ToString(testContext.DataRow["InternalMessage"]);
            string externalmessage = Convert.ToString(testContext.DataRow["ExternalMessage"]);

            pageMailboxDashboard.OpenAutomaticReply();
            string standing = pageMailboxDashboard.SetAutomaticReply(internalMessage,externalmessage);


            string status = VerifyResult(ExchangeMessages.AddAutomaticReply, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Automatic Reply", "Test to check if Automatic Reply is being added successfully", organizationName, "Mailbox", email, "Interanal Message: "+internalMessage+ "; External Message: "+externalmessage, status, standing);
            return status;
        }

        public string VerifyAutomaticReply(TestContext testContext)
        {
            string status = TestStatus.Success;
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string internalMessage = Convert.ToString(testContext.DataRow["InternalMessage"]);
            string externalmessage = Convert.ToString(testContext.DataRow["ExternalMessage"]);
                
            
            string standing = pageMailboxDashboard.VerifyAutomaticReply(internalMessage,externalmessage);

            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Automatic Reply", "Test to check if Automatic Reply been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;
        }

        public string AddArchive(TestContext testContext)
        {

            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            pageMailboxDashboard.OpenArchive();

                string standing = pageMailboxDashboard.SetArchive();


            string status = VerifyResult(ExchangeMessages.AddArchive, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Archive", "Test to check if archiving is being enabled or not", organizationName, "Mailbox", email, "", status, standing);
            if (status == TestStatus.Failed)
            {
                pageMailboxDashboard.CloseDialogueBox();
            }
            return status;
        }
        public string VerifyArchive(TestContext testContext)
        {
            string status = TestStatus.Success;
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string displayName = Convert.ToString(testContext.DataRow["NewDisplayName"]);


            string searchString;

            if (displayName == "")
            {
                MailAddress addr = new MailAddress(email);
                searchString = addr.User;
            }
            else
            {
                searchString = displayName;
            }

            string standing = pageMailboxDashboard.VerifyArchive(searchString);

            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify enabling of Archiving", "Test to check if archiving has been enabled or not", organizationName, "Mailbox", email, "", status, standing);
            return status;

        }

        public static void NavigateToMailboxDashboard(TestContext testContext)
        {
            try
            {

                //Arrange
                string email = Convert.ToString(testContext.DataRow["Email"]);
                string displayName = Convert.ToString(testContext.DataRow["DisplayName"]);


                ExgOrgMailboxes orgMailboxes = new ExgOrgMailboxes();
                orgMailboxes.SearchMailboxName(email, displayName);
                ExgMailboxDashboard mailboxDashboard = orgMailboxes.OpenMailboxDashboard();
               
                //Act
                
            }

            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);
            }

        }

        public static void NavigateToMailboxPage(TestContext testContext)
        {
            
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            try
            {
                SetDriverTime(2);
                if (DriverContext.Driver.FindElements(By.XPath("//h2")).Count == 0)
                {
                    SetDriverTime(30);
                    HomePage home = new HomePage();
                    home.ClickProvisioning();
                    ExchangeHome exgHome = home.ClickExchangeHome();
                    exgHome.SearchOrganizationName(organizationName);
                    ExgOrgMailboxes orgMailboxes = exgHome.MailboxesHome();
                }

                else if (DriverContext.Driver.FindElement(By.XPath("//h2")).Text != "Manage Mailboxes" && !DriverContext.Driver.FindElement(By.XPath("//p")).Text.Contains(organizationName))
                {
                    SetDriverTime(30);
                    PageRefresh(DriverContext.Driver);
                    
                    //Act
                    HomePage home = new HomePage();
                    home.ClickProvisioning();
                    ExchangeHome exgHome = home.ClickExchangeHome();
                    exgHome.SearchOrganizationName(organizationName);
                    ExgOrgMailboxes orgMailboxes = exgHome.MailboxesHome();
                }
                

            }

            catch (Exception )
            {
                
            }
            
        }

        public  void ClickMailboxBreakCrumb()
        {
            lnkMailboxes.ClickWithWait("header");
        }




        //public Tuple<IDictionary<string, string>, IDictionary<string, string>> VerifyUpdateMailboxGeneralProperties(
        //    TestContext testContext)
        //{
        //    try
        //    {
        //        btnTabRefreshButtonElem.Click();
        //        string firstname = Convert.ToString(testContext.DataRow["FirstName"]);
        //        string lastName = Convert.ToString(testContext.DataRow["LastName"]);
        //        string displayName = Convert.ToString(testContext.DataRow["NewDisplayName"]);
        //        string country = Convert.ToString(testContext.DataRow["Country"]);
        //        string state = Convert.ToString(testContext.DataRow["State"]);
        //        string officeLocation = Convert.ToString(testContext.DataRow["OfficeLocation"]);
        //        string address = Convert.ToString(testContext.DataRow["Address"]);
        //        string city = Convert.ToString(testContext.DataRow["City"]);
        //        string zipCode = Convert.ToString(testContext.DataRow["ZipCode"]);
        //        string jobTitle = Convert.ToString(testContext.DataRow["JobTitle"]);
        //        string company = Convert.ToString(testContext.DataRow["Company"]);
        //        string department = Convert.ToString(testContext.DataRow["Department"]);
        //        string managedBy = Convert.ToString(testContext.DataRow["ManagedBy"]);
        //        string businessPhone = Convert.ToString(testContext.DataRow["BusinessPhone"]);
        //        string fax = Convert.ToString(testContext.DataRow["Fax"]);
        //        string homePhone = Convert.ToString(testContext.DataRow["HomePhone"]);
        //        string mobilePhone = Convert.ToString(testContext.DataRow["MobilePhone"]);
        //        string pager = Convert.ToString(testContext.DataRow["Pager"]);
        //        string notes = Convert.ToString(testContext.DataRow["Notes"]);


        //        //Act
        //        IDictionary<string, string> expectedMailboxGeneralProperties = new Dictionary<string, string>();
        //        expectedMailboxGeneralProperties.Add("FirstName", firstname);
        //        expectedMailboxGeneralProperties.Add("LastName", lastName);
        //        expectedMailboxGeneralProperties.Add("DisplayName", displayName);
        //        expectedMailboxGeneralProperties.Add("Country", country);
        //        expectedMailboxGeneralProperties.Add("State", state);
        //        expectedMailboxGeneralProperties.Add("Office Location", officeLocation);
        //        expectedMailboxGeneralProperties.Add("Address", address);
        //        expectedMailboxGeneralProperties.Add("City", city);
        //        expectedMailboxGeneralProperties.Add("Zip Code", zipCode);
        //        expectedMailboxGeneralProperties.Add("Job Title", jobTitle);
        //        expectedMailboxGeneralProperties.Add("Company", company);
        //        expectedMailboxGeneralProperties.Add("Department", department);
        //        expectedMailboxGeneralProperties.Add("ManagedBy", managedBy);
        //        expectedMailboxGeneralProperties.Add("Business Phone", businessPhone);
        //        expectedMailboxGeneralProperties.Add("Fax", fax);
        //        expectedMailboxGeneralProperties.Add("Home Phone", homePhone);
        //        expectedMailboxGeneralProperties.Add("Mobile Phone", mobilePhone);
        //        expectedMailboxGeneralProperties.Add("Pager", pager);
        //        expectedMailboxGeneralProperties.Add("Notes", notes);




        //        ExgMailboxDashboard mailboxDashboard = new ExgMailboxDashboard();
        //        IDictionary<string, string> actualMailboxStorageProperties = mailboxDashboard.VerifyGeneralProperties();
        //        return Tuple.Create(expectedMailboxGeneralProperties, actualMailboxStorageProperties);
        //    }
        //    catch (Exception ex)
        //    {
        //        LogClass.AppendLogs(ex);
        //        return null;
        //    }
        //}


        //public Tuple<IDictionary<string, string>, IDictionary<string, string>> VerifyMailBoxStorageQuota(
        //    TestContext testContext)
        //{
        //    try
        //    {
        //        //Arrange
        //        string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
        //        string email = Convert.ToString(testContext.DataRow["Email"]);
        //        string mailboxSize = Convert.ToString(testContext.DataRow["MailboxSize"]);
        //        bool isCR = System.Convert.ToBoolean(testContext.DataRow["IsCR"]);
        //        string displayName = Convert.ToString(testContext.DataRow["DisplayName"]);
        //        //Act
        //        IDictionary<string, string> expectedMailboxStorageProperties = new Dictionary<string, string>();
        //        expectedMailboxStorageProperties.Add("MailboxSize", mailboxSize);
        //        expectedMailboxStorageProperties.Add("IncomingSize", mailboxSize);
        //        expectedMailboxStorageProperties.Add("OutgoingSize", mailboxSize);
        //        expectedMailboxStorageProperties.Add("ProhibitSendAt", mailboxSize);
        //        expectedMailboxStorageProperties.Add("IssueWarningAt", mailboxSize);
        //        expectedMailboxStorageProperties.Add("Quota", isCR == false ? "Accumulated" : mailboxSize);



        //        HomePage home = new HomePage();
        //        home.ClickProvisioning();
        //        ExchangeHome exgHome = home.ClickExchangeHome();
        //        exgHome.SearchOrganizationName(organizationName);
        //        ExgOrgMailboxes orgMailboxes = exgHome.MailboxesHome();
        //        orgMailboxes.SearchMailboxName(email, displayName);
        //        ExgMailboxDashboard mailboxDashboard = orgMailboxes.OpenMailboxDashboard();
        //        mailboxDashboard.OpenAdvancedProperties();
        //        IDictionary<string, string> actualMailboxStorageProperties = mailboxDashboard.GetStorageQuota();
        //        return Tuple.Create(expectedMailboxStorageProperties, actualMailboxStorageProperties);

        //    }

        //    catch (Exception ex)
        //    {
        //        LogClass.AppendLogs(ex);
        //        return null;

        //    }

        //}

    }
}

