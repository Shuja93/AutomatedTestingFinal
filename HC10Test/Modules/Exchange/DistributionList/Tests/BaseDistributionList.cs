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
using HC10Test.PageObject;
using OpenQA.Selenium;

namespace HC10Test
{
    public class BaseDistributionList : BasePage
    {
        private readonly ExgDLDashboard pageDlDashboard;

        public BaseDistributionList()
        {
            pageDlDashboard = new ExgDLDashboard();
        }

        public string CreateDl(TestContext testContext)
        {
            try
            {
                string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
                bool isSubOU = Convert.ToBoolean(testContext.DataRow["IsSubOU"]);
                string email = Convert.ToString(testContext.DataRow["Email"]);
                string groupType = Convert.ToString(testContext.DataRow["GroupType"]);
                bool isNewGroup = Convert.ToBoolean(testContext.DataRow["IsNewGroup"]);
                string adGroupName = Convert.ToString(testContext.DataRow["ADGroupName"]);
                string IncomingMessageSize = Convert.ToString(testContext.DataRow["IncomingMessageSize"]);
                bool isCR = Convert.ToBoolean(testContext.DataRow["IsCR"]);
                string displayName = Convert.ToString(testContext.DataRow["DisplayName"]);
                string adminUsers = Convert.ToString(testContext.DataRow["AdminUsers"]);
                string memberUsers = Convert.ToString(testContext.DataRow["MemberUsers"]);
                bool allSendersAuth = Convert.ToBoolean(testContext.DataRow["AllSendersAuth"]);



                //Act
                ExgOrgDL pageOrgDL = new ExgOrgDL();
                ExgCreateDistributionList pageCreateDL = pageOrgDL.OpenCreateDLPage();
                string standing = pageCreateDL.CreateDl(groupType, isSubOU, isNewGroup, adGroupName, email,
                    IncomingMessageSize, isCR, displayName, adminUsers,
                    memberUsers, allSendersAuth);


                //Verify
                var status = VerifyResult(ExchangeMessages.CreateDL, standing);
                if (status == TestStatus.Failed)
                {
                    pageOrgDL.CloseDialogueBox();
                }
                else
                {
                    Thread.Sleep(5000);
                }

                ReporterClass.Reporter("Exchange", Settings.UserLevel, "Create Exchange Group", "Exchange Group Creation Test",
                    organizationName, groupType, email,
                    "SubOU: " + isSubOU + "; IsNewGroup: " + isNewGroup + "; IsCr: " + isCR +
                    "; Incoming Message Size/CR Size :" + IncomingMessageSize, status, standing);
                TestTracker.distributionListStatus.Add(email, status);
                return status;
            }
            catch (Exception e)
            {
                LogClass.AppendLogs(e.Message);
                return TestStatus.Failed;
            }


        }

        public string AddAdvanceProperties(TestContext testContext)
        {
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string newDisplayName = Convert.ToString(testContext.DataRow["NewDisplayName"]); 
            bool hideFromAddessList = Convert.ToBoolean(testContext.DataRow["HideFromAddessList"]);
            bool sendOofMessageToOriginator = Convert.ToBoolean(testContext.DataRow["SendOOFMessagetoOriginator"]);
            bool SendersInsideandOutsideOrg = Convert.ToBoolean(testContext.DataRow["SendersInsideandOutsideOrg"]);
            string deliveryReport = Convert.ToString(testContext.DataRow["DeliveryReport"]);
            string notes = Convert.ToString(testContext.DataRow["Notes"]);
            string groupType = Convert.ToString(testContext.DataRow["GroupType"]);
            bool IsCR = Convert.ToBoolean(testContext.DataRow["IsCR"]);
            string incomingMessageSize = Convert.ToString(testContext.DataRow["IncomingMessage"]);



            string standing = pageDlDashboard.SetAdvanceProperties( newDisplayName,  hideFromAddessList,
             sendOofMessageToOriginator, SendersInsideandOutsideOrg,  deliveryReport,  notes, incomingMessageSize, IsCR);

            //Verify
            string status = VerifyResult(ExchangeMessages.UpdateDlAdvanceProperties, standing);
            ReporterClass.Reporter("Exchange", Settings.UserLevel, "Update Advance Properties of Exchange Group", "Test to check if Advance Properties of Exchange Group are updating properly or not", organizationName, groupType, email, "Hide from Address List", status, standing);

            return status;

        }

        public string AddMembersDL(TestContext testContext)
        {
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string groupType = Convert.ToString(testContext.DataRow["GroupType"]);
            string userList = Convert.ToString(testContext.DataRow["MemberUsers"]);

            pageDlDashboard.OpenMembers();
            string standing = pageDlDashboard.SetDlMembers(userList);

            //Verify
            string status = VerifyResult(ExchangeMessages.AddDlMembers, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Send On Behalf Users", "Test to check if Send On Behalf users are added successfully", organizationName, "Mailbox", email, "Email List: " + userList, status, standing);

            return status;

        }

        public string VerifyMembersDL(TestContext testContext)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string groupType = Convert.ToString(testContext.DataRow["GroupType"]);
            string userList = Convert.ToString(testContext.DataRow["MemberUsers"]);

            //Act
            pageDlDashboard.OpenMembers();
            string standing = pageDlDashboard.VerifyDlMembers(userList);

            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of new Send On Behalf users", "Test to check if the Send On Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);

            return status;
        }

        public string AddAdministratorDL(TestContext testContext)
        {
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string groupType = Convert.ToString(testContext.DataRow["GroupType"]);
            string userList = Convert.ToString(testContext.DataRow["AdminUsers"]);

            pageDlDashboard.OpenAdministrators();
            string standing = pageDlDashboard.SetDlAdmministrators(userList);

            //Verify
            string status = VerifyResult(ExchangeMessages.AddDlAdmins, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Send On Behalf Users", "Test to check if Send On Behalf users are added successfully", organizationName, "Mailbox", email, "Email List: " + userList, status, standing);

            return status;

        }

        public string VerifyAdministratorDL(TestContext testContext)
        {
            string status = TestStatus.Success;
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string groupType = Convert.ToString(testContext.DataRow["GroupType"]);
            string userList = Convert.ToString(testContext.DataRow["AdminUsers"]);

            //Act
            pageDlDashboard.OpenAdministrators();
            string standing = pageDlDashboard.VerifyDlAdministrators(userList);

            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of new Send On Behalf users", "Test to check if the Send On Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);

            return status;
        }



        public string AddAdditionalEmailAddress(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string newEmail = Convert.ToString(testContext.DataRow["AdditionalEmailAddress"]);
            string groupType = Convert.ToString(testContext.DataRow["GroupType"]);

            //Act
            pageDlDashboard.OpenEmailAddress();
            string standing = pageDlDashboard.SetAdditionalEmailAddress(newEmail);



            //Verify
            string status = VerifyResult(ExchangeMessages.AddDlEmailAddress, standing);
            ReporterClass.Reporter("Exchange", "Host", "Add Email Address", "Test to check if email addresses are added as additional aliases or not", organizationName,groupType, email, "Email: " + newEmail, status, standing);

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
            string standing = pageDlDashboard.VerifyAdditionalEmailAddress(newEmail);


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
            pageDlDashboard.OpenSendOnBehalf();
            string standing = pageDlDashboard.SetSendOnBehalf(userList);

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
            string standing = pageDlDashboard.VerifySendOnBehalf(userList);

            //Verify
            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of new Send On Behalf users", "Test to check if the Send On Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);

            return status;
        }

        public string AddSendAsPermissions(TestContext testContext)
        {
            //Stage
            string organizationName = Convert.ToString(testContext.DataRow["OrganizationName"]);
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["SendAsPermissionsUsers"]);
            pageDlDashboard.OpenSendAsPermissions();

            //Act
            string standing = pageDlDashboard.SetSendAsPermissions(userList);

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
            string standing = pageDlDashboard.VerifySendAsPermissions(userList);


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
            pageDlDashboard.OpenAcceptedSenders();

            //Act
            string standing = pageDlDashboard.SetAcceptedSenders(userList);

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
            string standing = pageDlDashboard.VerifyAcceptedSenders(userList);

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
            string email = Convert.ToString(testContext.DataRow["Email"]);
            string userList = Convert.ToString(testContext.DataRow["RejectedSenders"]);
            pageDlDashboard.OpenRejectedSenders();

            //Act
            string standing = pageDlDashboard.SetRejectedSenders(userList);

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

            string standing = pageDlDashboard.VerifyRejectedSenders(userList);

            if (standing != TestStatus.Success)
            {
                status = TestStatus.Failed;
            }
            ReporterClass.Reporter("Exchange", "Host", "Verify Addition of Rejected Users", "Test to check if Rejected Users have been assigned correctly or not", organizationName, "Mailbox", email, "", status, standing);
            return status;


        }

        public void ClickDlBreakCrumb()
        {
            lnkDl.ClickWithWait("header");
        }








        public static void NavigateToDLDashboard(TestContext testContext)
        {
            try
            {

                //Arrange
                string email = Convert.ToString(testContext.DataRow["Email"]);
                string displayName = Convert.ToString(testContext.DataRow["DisplayName"]);


                ExgOrgDL orgDL = new ExgOrgDL();
                orgDL.SearchDL(email, displayName);
                ExgDLDashboard mailboxDashboard = orgDL.OpenDLDashboard();

                //Act

            }

            catch (Exception ex)
            {
                LogClass.AppendLogs(ex);
            }

        }

        public static void NavigateToDlPage(TestContext testContext)
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
                    exgHome.DistributionListsHome();
                }

                else if (DriverContext.Driver.FindElement(By.XPath("//h2")).Text != "Manage Distribution Lists" && !DriverContext.Driver.FindElement(By.XPath("//p")).Text.Contains(organizationName))
                {
                    SetDriverTime(30);
                    PageRefresh(DriverContext.Driver);

                    //Act
                    HomePage home = new HomePage();
                    home.ClickProvisioning();
                    ExchangeHome exgHome = home.ClickExchangeHome();
                    exgHome.SearchOrganizationName(organizationName);
                    exgHome.DistributionListsHome();
                }


                
            }

            catch (Exception)
            {

            }

        }
    }
}
