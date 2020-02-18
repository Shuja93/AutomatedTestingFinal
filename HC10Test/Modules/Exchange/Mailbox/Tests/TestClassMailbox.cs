using System;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Config;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.TestTracker;
using System.Web.Hosting;

namespace HC10Test
{

    [TestClass]
    public class TestClassMailbox : BaseMailbox
    {
        public TestContext TestContext { get; set; }
        private SoftAssertions _softAssertions;
        private readonly string userLevel;

        public TestClassMailbox()
        {
            userLevel = Settings.UserLevel.ToLower();
        }


        [ClassInitialize]
        public static void ClassSetup(TestContext TestContext)
        {
            var testInitialize = new TestInitialize();
            testInitialize.InitializeSettings();
            
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            DriverContext.Driver.Close();
            DriverContext.Driver.Quit();
        }

        [TestInitialize]
        public void SetUp()
        {
            _softAssertions = new SoftAssertions();
        }

        [TestCleanup]
        public void TearDown()
        {
            _softAssertions.AssertAll();
        }

       
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\Mailbox\\Data\\MailboxCreation.csv", "MailboxCreation#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        [TestCategory("Mailbox")]
    
        
        public void MailboxCreation()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != userLevel)
            {
                Assert.Inconclusive();
            }

            NavigateToMailboxPage(TestContext);
            _softAssertions.Add("Test Create Mailbox", TestStatus.Success, CreateMailbox(TestContext));

            
            if (TestTracker.mailboxStatus[Convert.ToString(TestContext.DataRow["Email"])] == TestStatus.Success)
            {
                NavigateToMailboxDashboard(TestContext);
                _softAssertions.Add("Test Verify Mailbox General Properties", TestStatus.Success, VerifyMailBoxGeneralProfile(TestContext,true));
                _softAssertions.Add("Test Verify Mailbox Advance Properties", TestStatus.Success, VerifyMailBoxAdvanceProperties(TestContext,true));
                ClickMailboxBreakCrumb();
            }
        }


        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\Mailbox\\Data\\MailboxUpdate.csv", "MailboxUpdate#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        [TestCategory("Mailbox")]
        
        public void MailboxUpdateDashboard()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != userLevel)
            {
                Assert.Inconclusive();
            }

            if (TestTracker.mailboxStatus[Convert.ToString(TestContext.DataRow["Email"])] == TestStatus.Success)
            { 
                NavigateToMailboxPage(TestContext);
                NavigateToMailboxDashboard(TestContext);

                _softAssertions.Add("Test Update General Properties", TestStatus.Success, UpdateMailboxGeneralProperties(TestContext));
                _softAssertions.Add("Test Verify Update General Properties", TestStatus.Success, VerifyMailBoxGeneralProfile(TestContext,false));

                _softAssertions.Add("Test  Update Advance Properties", TestStatus.Success, UpdateMailboxAdvanceProperties(TestContext));
                _softAssertions.Add("Test  Update Advance Properties", TestStatus.Success, VerifyMailBoxAdvanceProperties(TestContext,false));

                _softAssertions.Add("Test Update Retention Policy", TestStatus.Success, UpdateRetentionPolicy(TestContext));
                _softAssertions.Add("Test Verify Update Retention Policy", TestStatus.Success, VerifyRetentionPolicy(TestContext));

                _softAssertions.Add("Test Add Email Address", TestStatus.Success, AddAdditionalEmailAddress(TestContext));
                _softAssertions.Add("Test Verify Add Email Address", TestStatus.Success, VerifyAdditionalEmailAddress(TestContext));

                _softAssertions.Add("Test Add SendOnBehalf Users", TestStatus.Success, AddSendOnBehalfUsers(TestContext));
                _softAssertions.Add("Test Verify Add SendOnBehalf Users", TestStatus.Success, VerifyAddSendOnBehalfUsers(TestContext));

                _softAssertions.Add("Test Add FullAccessPermissions Users", TestStatus.Success, AddFullAccessPermissions(TestContext));
                _softAssertions.Add("Test Verify Add FullAccessPermissions Users", TestStatus.Success, VerifyFullAccessPermissions(TestContext));

                _softAssertions.Add("Test Add SendAsPermissions Users", TestStatus.Success, AddSendAsPermissions(TestContext));
                _softAssertions.Add("Test Verify Add SendAsPermissions Users", TestStatus.Success, VerifySendAsPermissions(TestContext));

                _softAssertions.Add("Test Add AcceptedSenders Users", TestStatus.Success, AddAcceptedSenders(TestContext));
                _softAssertions.Add("Test Verify Add AcceptedSenders Users", TestStatus.Success, VerifyAcceptedSenders(TestContext));

                _softAssertions.Add("Test Add Rejected Users", TestStatus.Success, AddRejectedSenders(TestContext));
                _softAssertions.Add("Test Verify Add Rejected Users", TestStatus.Success, VerifyRejectedSenders(TestContext));

                _softAssertions.Add("Test Add Forwarding User", TestStatus.Success, AddForwarding(TestContext));
                _softAssertions.Add("Test Verify Add Forwarding Users", TestStatus.Success, VerifyForwarding(TestContext));

                _softAssertions.Add("Test Add Automatic Reply", TestStatus.Success, AddAutomaticReply(TestContext));
               // _softAssertions.Add("Test Verify Add Automatic Replys", TestStatus.Success, VerifyAutomaticReply(TestContext));

                _softAssertions.Add("Test Add Archive", TestStatus.Success, AddArchive(TestContext));
                //_softAssertions.Add("Test Verify Add Archive", TestStatus.Success, VerifyArchive(TestContext));
                ClickMailboxBreakCrumb();

            }
        }
    }
}
