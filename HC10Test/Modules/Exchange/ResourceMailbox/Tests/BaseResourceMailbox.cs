using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Logs;
using HC10AutomationFramework.Enum;
using System.Net.Mail;
using System.Threading;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Helpers;
using HC10AutomationFramework.TestTracker;
using OpenQA.Selenium;

namespace HC10Test
{
     public class BaseResourceMailbox : BasePage

     {
         private IWebElement btnAddForwardingUser => ByXPath("//*[@id='btn-rmbuser']");
         private readonly ExgResourceMailboxDashboard pageResourceMailboxDashboard;

        public BaseResourceMailbox()
        {
            pageResourceMailboxDashboard = new ExgResourceMailboxDashboard();
        }
        public string CreateResourceMailbox(TestContext testContext)
        {
            try
            {
                //revisit - VerifyOUMethod
                //Stage
                string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
                string mailboxType = Convert.ToString(testContext.DataRow["MailboxType"]);
                bool isSubOU = Convert.ToBoolean(testContext.DataRow["IsSubOU"]);
                string email = Convert.ToString(testContext.DataRow["Email"]);
                string mailboxSize = Convert.ToString(testContext.DataRow["MailboxSize"]);
                bool isCR = Convert.ToBoolean(testContext.DataRow["IsCR"]);
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
                string businessPhone = Convert.ToString(testContext.DataRow["BusinessPhone"]);
                string fax = Convert.ToString(testContext.DataRow["Fax"]);
                string homePhone = Convert.ToString(testContext.DataRow["HomePhone"]);
                string mobilePhone = Convert.ToString(testContext.DataRow["MobilePhone"]);
                string pager = Convert.ToString(testContext.DataRow["Pager"]);
                string notes = Convert.ToString(testContext.DataRow["Notes"]);


                //Act
                ExgOrgResourceMailboxes pageOrgResourceMailboxes = new ExgOrgResourceMailboxes();
                ExgCreateResourceMailbox pageCreateResourceMailbox = pageOrgResourceMailboxes.OpenCreateResourceMailboxPage();
                string standing = pageCreateResourceMailbox.CreateResourceMailbox(mailboxType, isSubOU, email, isCR,
                    mailboxSize, firstname, lastName, displayName, country, state, officeLocation, address, city, zipCode, jobTitle, company, department, businessPhone,
                    fax, homePhone, mobilePhone, pager, notes);


                //Verify
                var status = VerifyResult(ExchangeMessages.CreateResourceMailbox, standing);
                if (status == TestStatus.Failed)
                {
                   CloseDialogueBox();
                }
                else
                {
                    Thread.Sleep(5000);
                }
                //ReporterClass.Reporter("Exchange", "Host", "Create Mailbox", "Mailbox Creation Test", organizationName, "Mailbox", email, "SubOU: " + isSubOU + "; IsNewUser: " + isNewUser + "; IsCr: " + isCR + "; Mailbox/CR Size :" + mailboxSize, status, standing);
                TestTracker.resourceMailboxStatus.Add(email, status);

                return status;
            }
            catch (Exception e)
            {
                LogClass.AppendLogs(e.Message);
                return e.Message;
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
            string standing = pageResourceMailboxDashboard.VerifyGeneralProperties(firstname, lastName, displayName, country, state,
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

            if (isNewMailbox)
            {
                pageResourceMailboxDashboard.OpenAdvancedProperties();
                
            }
            btnTabRefreshButtonElem.ClickWithWait("spinner");
            
            string standing = pageResourceMailboxDashboard.VerifyAdvanceProperties(mailboxSize,isCR);
        

    
            //Verify
            if (!string.IsNullOrWhiteSpace(standing))
            {
                status = TestStatus.Failed;
            }

            if (isNewMailbox)
            {
                ReporterClass.Reporter("Exchange", "Host", "Verify New mailbox Advance Properties", "Test to verify that the Advance Properties set at the time of mailbox creation are set successfully", organizationName, "Mailbox", email, "", status, standing);

            }
            else
            {
                ReporterClass.Reporter("Exchange", "Host", "Verify New mailbox Advance Properties", "Test to verify that the Advance Properties set at the time of mailbox update are set successfully", organizationName, "Mailbox", email, "", status, standing);
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
            string standing = pageResourceMailboxDashboard.SetGeneralProperties(firstname, lastName, displayName,country, state,
                officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
                fax, homePhone, mobilePhone, pager, notes);
        

      
            //Verify
            string status = VerifyResult(ExchangeMessages.UpdateMailboxGeneralProperties, standing);
            ReporterClass.Reporter("Exchange", "Host", "Update Mailbox General Properties", "Test to verify if General Properties are being updated properly or not", organizationName, "Mailbox", email, "Refer to CSV File", status, standing);
            return status;
        }
        
       

        public string AddAdditionalEmailAddress(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string newEmail = Convert.ToString(testContext.DataRow["AdditionalEmailAddress"]);

            //Act
            pageResourceMailboxDashboard.OpenEmailAddress();
            string standing = pageResourceMailboxDashboard.SetAdditionalEmailAddress(newEmail);
        
       
        
            //Verify
            string status = VerifyResult(ExchangeMessages.AddResourceEmailAddress, standing);
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
            string standing = pageResourceMailboxDashboard.VerifyAdditionalEmailAddress(newEmail);
        
       
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
            pageResourceMailboxDashboard.OpenSendOnBehalf();
            string standing = pageResourceMailboxDashboard.SetSendOnBehalf(userList);

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
            string standing = pageResourceMailboxDashboard.VerifySendOnBehalf(userList);
       
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
            pageResourceMailboxDashboard.OpenFullAccessPermissions();
            string standing = pageResourceMailboxDashboard.SetFullAccessPermissions(userList);
        

        
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
            string standing = pageResourceMailboxDashboard.VerifyFullAccessPermissions(userList);
      
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
            pageResourceMailboxDashboard.OpenSendAsPermissions();
            
            //Act
            string standing = pageResourceMailboxDashboard.SetSendAsPermissions(userList);

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
            string standing = pageResourceMailboxDashboard.VerifySendAsPermissions(userList);


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
            pageResourceMailboxDashboard.OpenAcceptedSenders();

            //Act
            string standing = pageResourceMailboxDashboard.SetAcceptedSenders(userList);
            
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
            string standing = pageResourceMailboxDashboard.VerifyAcceptedSenders(userList);

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
            pageResourceMailboxDashboard.OpenRejectedSenders();

            //Act
            string standing = pageResourceMailboxDashboard.SetRejectedSenders(userList);

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

                string standing= pageResourceMailboxDashboard.VerifyRejectedSenders(userList);

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

            pageResourceMailboxDashboard.OpenForwarding();
            string standing = pageResourceMailboxDashboard.SetForwarding(user,ou,exchangeObject, btnAddForwardingUser);

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

            string standing = pageResourceMailboxDashboard.VerifyForwarding(user);

            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Forwarding User", "Test to check if forwarding Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;
        }

        
        public string AddArchive(TestContext testContext)
        {

            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            pageResourceMailboxDashboard.OpenArchive();

            string standing = pageResourceMailboxDashboard.SetArchive();


            string status = VerifyResult(ExchangeMessages.AddArchive, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Archive", "Test to check if archiving is being enabled or not", organizationName, "Mailbox", email, "", status, standing);
            pageResourceMailboxDashboard.CloseDialogueBox();
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

            string standing = pageResourceMailboxDashboard.VerifyArchive(searchString);

            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify enabling of Archiving", "Test to check if archiving has been enabled or not", organizationName, "Mailbox", email, "", status, standing);
            return status;

        }

        public static void NavigateToResourceMailboxDashboard(TestContext testContext)
        {
            try
            {

                //Arrange
                string email = Convert.ToString(testContext.DataRow["Email"]);
                string displayName = Convert.ToString(testContext.DataRow["DisplayName"]);


                ExgOrgResourceMailboxes orgResourceMailboxes = new ExgOrgResourceMailboxes();
                orgResourceMailboxes.SearchResourceMailboxName(email, displayName);
                ExgResourceMailboxDashboard resourceMailboxDashboard =
                    orgResourceMailboxes.OpenResourceMailboxDashboard();

                //Act

            }

            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);

            }

        }

        public static void NavigateToResourceMailboxPage(TestContext testContext)
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
                    ExgOrgResourceMailboxes orgResourceMailboxes = exgHome.ResourceMailboxHome();
                }

                else if (DriverContext.Driver.FindElement(By.XPath("//h2")).Text != "Manage Resource Mailboxes" && !DriverContext.Driver.FindElement(By.XPath("//p")).Text.Contains(organizationName))
                {
                    SetDriverTime(30);
                    PageRefresh(DriverContext.Driver);
                    
                    //Act
                    HomePage home = new HomePage();
                    home.ClickProvisioning();
                    ExchangeHome exgHome = home.ClickExchangeHome();
                    exgHome.SearchOrganizationName(organizationName);
                    ExgOrgResourceMailboxes orgResourceMailboxes = exgHome.ResourceMailboxHome();
                }
                

                SetDriverTime(30);
            }

            catch (Exception )
            {
                
            }
            
        }

        public  void ClickResourceMailboxBreakCrumb()
        {
            lnkResourceMailboxes.ClickWithWait("header");
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

