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
using HC10Test.PageObject;
using Microsoft.SqlServer.Server;
using OpenQA.Selenium;

namespace HC10Test
{
    public class BaseMailContact : BasePage
    {


        private readonly ExgMailContactDashboard pageMailContactDashboard;
        public BaseMailContact()
        {
            pageMailContactDashboard = new ExgMailContactDashboard();
        } 



        public string CreateMailContact(TestContext testContext)
        {
            try
            {
                //revisit - VerifyOUMethod
                //Stage
                string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
                bool isSubOU = Convert.ToBoolean(testContext.DataRow["IsSubOU"]);
                string internalEmail = Convert.ToString(testContext.DataRow["InternalEmailAddress"]);
                string externalEmail = Convert.ToString(testContext.DataRow["ExternalEmailAddress"]);
                string contactName = Convert.ToString(testContext.DataRow["ContactName"]);
                bool isHiddenFromAddressList = Convert.ToBoolean(testContext.DataRow["IsHiddenFromAddressList"]);
                string maximumRecipients = Convert.ToString(testContext.DataRow["MaximumRecipients"]);
                string maximumReceiveSize = Convert.ToString(testContext.DataRow["MaximumReceiveSize"]);
                string firstName = Convert.ToString(testContext.DataRow["FirstName"]);
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
                ExgOrgMailContacts pageMailContacts = new ExgOrgMailContacts();
                ExgCreateMailContact pagecreateMailContact = pageMailContacts.OpenCreateMailContactPage();
                string standing = pagecreateMailContact.CreateMailContact(isSubOU, contactName, firstName, lastName,
                    externalEmail, internalEmail, isHiddenFromAddressList, maximumRecipients,
                    maximumReceiveSize, displayName, country, state, officeLocation, address, city,
                    zipCode, jobTitle, company, department, managedBy, businessPhone,
                    fax, homePhone, mobilePhone, pager, notes);


                //Verify
                var status = VerifyResult(ExchangeMessages.CreateMailContact, standing);
                if (status == TestStatus.Failed)
                {
                    CloseDialogueBox();
                }
                else
                {
                    Thread.Sleep(5000);
                }

                //ReporterClass.Reporter("Exchange", "Host", "Create Mailbox", "Mailbox Creation Test", organizationName,
                //    "Mailbox", email,
                //    "SubOU: " + isSubOU + "; IsNewUser: " + isNewUser + "; IsCr: " + isCR + "; Mailbox/CR Size :" +
                //    mailboxSize, status, standing);

                if (string.IsNullOrEmpty(internalEmail))
                {
                    TestTracker.mailContactStatus.Add(externalEmail, status);
                }
                else
                {
                    TestTracker.mailContactStatus.Add(internalEmail, status);
                }
                
                return status;
            }
            catch (Exception e)
            {
                LogClass.AppendLogs(e.Message);
                return TestStatus.Failed;
            }
        }

        public string VerifyMailContactGeneralProfile(TestContext testContext, bool isNewMailbox)
        {
            //Stage

            string status = TestStatus.Success;
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string firstname = Convert.ToString(testContext.DataRow["FirstName"]);
            string lastName = Convert.ToString(testContext.DataRow["LastName"]);
            string displayName;
            if (isNewMailbox) { displayName = Convert.ToString(testContext.DataRow["DisplayName"]); }
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
            string standing = pageMailContactDashboard.VerifyGeneralProperties(firstname, lastName, displayName, country, state,
                officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
                fax, homePhone, mobilePhone, pager, notes);

            //Verify
            if (!string.IsNullOrEmpty(standing))
            {
                status = TestStatus.Failed;
            }

            if (isNewMailbox)
            {
                //ReporterClass.Reporter("Exchange", "Host", "Verify New mailbox General Properties", "Test to verify that the General Properties set at the time of mailbox creation are set successfully", organizationName, "Mailbox", email, "", status, standing);
            }
            else
            {
                //ReporterClass.Reporter("Exchange", "Host", "Verify New mailbox General Properties", "Test to verify that the General Properties set at the time of mailbox update are set successfully", organizationName, "Mailbox", email, "", status, standing);
            }
            return status;
        }

        public string UpdateMailContactGeneralProperties(TestContext testContext)
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
            string standing = pageMailContactDashboard.SetGeneralProperties(firstname, lastName, displayName, country, state,
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
            string email = Convert.ToString(testContext.DataRow["InternalEmailAddress"]);
            string newEmail = Convert.ToString(testContext.DataRow["AdditionalEmailAddress"]);

            //Act
            pageMailContactDashboard.OpenEmailAddress();
            string standing = pageMailContactDashboard.SetAdditionalEmailAddress(newEmail);



            //Verify
            string status = VerifyResult(ExchangeMessages.AddMailContactEmailAddress, standing);
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
            string standing = pageMailContactDashboard.VerifyAdditionalEmailAddress(newEmail);


            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of new email alias", "Test to check if the email address has been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;
        }


        public string AddAcceptedSenders(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["InternalEmailAddress"]);
            string userList = Convert.ToString(testContext.DataRow["AcceptedSenders"]);
            pageMailContactDashboard.OpenAcceptedSenders();

            //Act
            string standing = pageMailContactDashboard.SetAcceptedSenders(userList);

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
            string standing = pageMailContactDashboard.VerifyAcceptedSenders(userList);

            //Verify
            if (standing != TestStatus.Success)
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
            string email = Convert.ToString(testContext.DataRow["InternalEmailAddress"]);
            string userList = Convert.ToString(testContext.DataRow["RejectedSenders"]);
            pageMailContactDashboard.OpenRejectedSenders();

            //Act
            string standing = pageMailContactDashboard.SetRejectedSenders(userList);

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

            string standing = pageMailContactDashboard.VerifyRejectedSenders(userList);

            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Rejected Users", "Test to check if Rejected Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;


        }

        public string VerifyExternalEmail(TestContext testContext)
        {
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string externalEmail = Convert.ToString(testContext.DataRow["ExternalEmailAddress"]);

            string status = pageMailContactDashboard.VerifyExternalEmailAddress(externalEmail);
            //ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Rejected Users", "Test to check if Rejected Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;
        }

        public string VerifyInternalEmail(TestContext testContext)
        {
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string internalEmail = Convert.ToString(testContext.DataRow["InternalEmailAddress"]);
            string externalEmail = Convert.ToString(testContext.DataRow["ExternalEmailAddress"]);

            string status = pageMailContactDashboard.VerifyInternalEmailAddress(externalEmail, internalEmail);
            //ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Rejected Users", "Test to check if Rejected Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;
        }






        public static void NavigateToMailContact(TestContext testContext)
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
                    exgHome.MailContactHome();
                }

                else if (DriverContext.Driver.FindElement(By.XPath("//h2")).Text != "Manage Mail Contacts" && !DriverContext.Driver.FindElement(By.XPath("//p")).Text.Contains(organizationName))
                {
                    SetDriverTime(30);
                    PageRefresh(DriverContext.Driver);

                    //Act
                    HomePage home = new HomePage();
                    home.ClickProvisioning();
                    ExchangeHome exgHome = home.ClickExchangeHome();
                    exgHome.SearchOrganizationName(organizationName);
                    exgHome.MailContactHome();
                }


                
            }

            catch (Exception)
            {

            }

        }

        public static void NavigateToMailContactDashboard(TestContext testContext)
        {
            try
            {

                //Arrange
                string internalEmail = Convert.ToString(testContext.DataRow["InternalEmailAddress"]);
                string contactName = Convert.ToString(testContext.DataRow["DisplayName"]);


                ExgOrgMailContacts orgMailContacts = new ExgOrgMailContacts();
                orgMailContacts.SearchMailContact(internalEmail, contactName);
                ExgMailContactDashboard mailContactDashboard = orgMailContacts.OpenMailContactDashboard();

                //Act

            }
            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);
            }
        }

        public void ClickMailContactrumb()
        {
            lnkMailContact.ClickWithWait("header");
        }
    }
}
