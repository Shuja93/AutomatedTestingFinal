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
    [TestCategory("PublicFolder")]
    public class TestClassPublicFolder : BasePublicFolder
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

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\PublicFolder\\Data\\PublicFolderCreate.csv", "PublicFolderCreate#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]

        public void PublicFolderCreation()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != Settings.UserLevel.ToLower())
            {
                Assert.Inconclusive();
            }


            NavigateToPublicFolderPage(TestContext);
            _softAssertions.Add("Test Create Mailbox", TestStatus.Success, CreatePublicFolder(TestContext));
        }

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\PublicFolder\\Data\\PublicFolderDashboard.csv", "PublicFolderDashboard#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        
        public void PublicFolderUpdateDashboard()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != Settings.UserLevel.ToLower())
            {
                Assert.Inconclusive();
            }

            if (TestTracker.publicFolderStatus[Convert.ToString(TestContext.DataRow["Email"])] == TestStatus.Success)
            {
                NavigateToPublicFolderPage(TestContext);
                NavigateToPublicFolderDashboard(TestContext);

                _softAssertions.Add("Test Add Email Address", TestStatus.Success, AddAdditionalEmailAddress(TestContext));
                _softAssertions.Add("Test Verify Add Email Address", TestStatus.Success, VerifyAdditionalEmailAddress(TestContext));

                _softAssertions.Add("Test Add AcceptedSenders Users", TestStatus.Success, AddAcceptedSenders(TestContext));
                _softAssertions.Add("Test Verify Add AcceptedSenders Users", TestStatus.Success, VerifyAcceptedSenders(TestContext));

                _softAssertions.Add("Test Add Rejected Users", TestStatus.Success, AddRejectedSenders(TestContext));
                _softAssertions.Add("Test Verify Add SendOnBehalf Users", TestStatus.Success, VerifyRejectedSenders(TestContext));

                _softAssertions.Add("Test Add Forwarding User", TestStatus.Success, AddForwarding(TestContext));
                _softAssertions.Add("Test Verify Add Forwarding Users", TestStatus.Success, VerifyForwarding(TestContext));

                ClickPublicFolderCrumb();

            }


        }
    }
}
