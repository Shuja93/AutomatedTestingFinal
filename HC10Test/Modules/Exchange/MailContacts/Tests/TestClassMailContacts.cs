using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Config;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Logs;
using HC10AutomationFramework.TestTracker;

namespace HC10Test
{
    [TestClass]
    [TestCategory("MailContact")]
    public class TestClassMailContacts : BaseMailContact
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

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\MailContacts\\Data\\MailContactCreate.csv", "MailContactCreate#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        
        public void MailContactCreation()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != Settings.UserLevel.ToLower())
            {
                Assert.Inconclusive();
            }
            NavigateToMailContact(TestContext);
            _softAssertions.Add("Test Create Mailbox", TestStatus.Success, CreateMailContact(TestContext));


            if (TestTracker.mailContactStatus[Convert.ToString(TestContext.DataRow["ExternalEmailAddress"])] == TestStatus.Success)
            {
                NavigateToMailContactDashboard(TestContext);
                _softAssertions.Add("Test Verify Mailbox General Properties", TestStatus.Success, VerifyMailContactGeneralProfile(TestContext, true));
                _softAssertions.Add("Test Verify External Email", TestStatus.Success, VerifyExternalEmail(TestContext));
                _softAssertions.Add("Test Verify External Email", TestStatus.Success, VerifyInternalEmail(TestContext));
                ClickMailContactrumb();
               
            }
        }

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\MailContacts\\Data\\MailContactDashboard.csv", "MailContactDashboard#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]

        public void MailContactUpdateDashboard()
        {
            try
            {
                if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != Settings.UserLevel.ToLower())
                {
                    Assert.Inconclusive();
                }

                if (TestTracker.mailContactStatus[Convert.ToString(TestContext.DataRow["ExternalEmailAddress"])] == TestStatus.Success)
                {
                    NavigateToMailContact(TestContext);
                    NavigateToMailContactDashboard(TestContext);


                    _softAssertions.Add("Test Add General Properties", TestStatus.Success, UpdateMailContactGeneralProperties(TestContext));
                    _softAssertions.Add("Test Verify Add General Properties", TestStatus.Success, VerifyMailContactGeneralProfile(TestContext, false));

                    _softAssertions.Add("Test Add Advance Properties", TestStatus.Success, AddAdvanceProperties(TestContext));

                    _softAssertions.Add("Test Add Email Address", TestStatus.Success, AddAdditionalEmailAddress(TestContext));
                    _softAssertions.Add("Test Verify Email Address", TestStatus.Success, VerifyAdditionalEmailAddress(TestContext));


                    _softAssertions.Add("Test Add AcceptedSenders Users", TestStatus.Success, AddAcceptedSenders(TestContext));
                    _softAssertions.Add("Test Verify Add AcceptedSenders Users", TestStatus.Success, VerifyAcceptedSenders(TestContext));


                    _softAssertions.Add("Test Add Rejected Users", TestStatus.Success, AddRejectedSenders(TestContext));
                    _softAssertions.Add("Test Verify Add Rejected Users", TestStatus.Success, VerifyRejectedSenders(TestContext));

                    ClickMailContactrumb();

                }

            }
            catch (Exception e)
            {
                LogClass.AppendLogs(e.Message);
                throw;
            }
           


        }
    }
}
