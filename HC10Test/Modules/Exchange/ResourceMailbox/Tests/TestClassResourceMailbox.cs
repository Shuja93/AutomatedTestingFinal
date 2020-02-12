using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Config;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.TestTracker;


namespace HC10Test
{

    [TestClass]
    [TestCategory("ResourceMailbox")]
    public class TestClassResourceMailbox : BaseResourceMailbox
    {
        public TestContext TestContext { get; set; }
        private SoftAssertions _softAssertions;

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

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\ResourceMailbox\\Data\\ResourceMailboxCreate.csv", "ResourceMailboxCreate#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
       
        public void ResourceMailboxCreation()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != Settings.UserLevel.ToLower())
            {
                Assert.Inconclusive();
            }

            NavigateToResourceMailboxPage(TestContext);
            _softAssertions.Add("Test Create Mailbox", TestStatus.Success, CreateResourceMailbox(TestContext));

            if (TestTracker.resourceMailboxStatus[Convert.ToString(TestContext.DataRow["Email"])] == TestStatus.Success)
            {
                NavigateToResourceMailboxDashboard(TestContext);
                _softAssertions.Add("Test Verify Mailbox General Properties", TestStatus.Success, VerifyMailBoxGeneralProfile(TestContext,true));
                _softAssertions.Add("Test Verify Mailbox Advance Properties", TestStatus.Success, VerifyMailBoxAdvanceProperties(TestContext,true));
                ClickResourceMailboxBreakCrumb();
            }

           
        }




        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\ResourceMailbox\\Data\\ResourceMailboxDashboard.csv", "ResourceMailboxDashboard#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
       
        public void ResourceMailboxUpdateDashboard()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != Settings.UserLevel.ToLower())
            {
                Assert.Inconclusive();
            }

            if (TestTracker.resourceMailboxStatus[Convert.ToString(TestContext.DataRow["Email"])] == TestStatus.Success)
            { 
                NavigateToResourceMailboxPage(TestContext);
                NavigateToResourceMailboxDashboard(TestContext);

                _softAssertions.Add("Test Update General Properties", TestStatus.Success, UpdateMailboxGeneralProperties(TestContext));
                _softAssertions.Add("Test Verify Update General Properties", TestStatus.Success, VerifyMailBoxGeneralProfile(TestContext,false));

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
                _softAssertions.Add("Test Verify Add SendOnBehalf Users", TestStatus.Success, VerifyRejectedSenders(TestContext));

                _softAssertions.Add("Test Add Forwarding User", TestStatus.Success, AddForwarding(TestContext));
                _softAssertions.Add("Test Verify Add Forwarding Users", TestStatus.Success, VerifyForwarding(TestContext));

               //_softAssertions.Add("Test Add Archive", TestStatus.Success, AddArchive(TestContext));
               //_softAssertions.Add("Test Verify Add Archive", TestStatus.Success, VerifyArchive(TestContext));
                ClickResourceMailboxBreakCrumb();

            }
        }
    }




}
