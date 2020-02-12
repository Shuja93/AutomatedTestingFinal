using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Config;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.TestTracker;
using HC10Test.PageObjects;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HC10Test
{
    [TestClass]
    public class TestClassDL : BaseDistributionList
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

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\Modules\\Exchange\\DistributionList\\Data\\DLCreation.csv", "DLCreation#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        [TestCategory("DistributionLists")]
        public void CreateDistributionLists()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != Settings.UserLevel.ToLower())
            {
                Assert.Inconclusive();
            }

            NavigateToDlPage(TestContext);
            _softAssertions.Add("Test Create DL", TestStatus.Success, CreateDl(TestContext));


            if (TestTracker.distributionListStatus[Convert.ToString(TestContext.DataRow["Email"])] == TestStatus.Success)
            {
                NavigateToDLDashboard(TestContext);
                _softAssertions.Add("Test Verify DL Members", TestStatus.Success, VerifyMembersDL(TestContext,true));
                _softAssertions.Add("Test Verify Mailbox Advance Properties", TestStatus.Success, VerifyAdministratorDL(TestContext,true));
                ClickDlBreakCrumb();
            }
        }

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV","|DataDirectory|\\Modules\\Exchange\\DistributionList\\Data\\DLUpdate.csv", "DLUpdate#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        [TestCategory("DistributionLists")]
        public void DistributionListsUpdateDashboard()
        {
            if (Convert.ToString(TestContext.DataRow["Userlevel"]).ToLower() != Settings.UserLevel.ToLower())
            {
                Assert.Inconclusive();
            }

            if (TestTracker.distributionListStatus[Convert.ToString(TestContext.DataRow["Email"])] == TestStatus.Success)
            {
                NavigateToDlPage(TestContext);
                NavigateToDLDashboard(TestContext);

                _softAssertions.Add("Test Add Advance Properties", TestStatus.Success,
                    AddAdvanceProperties(TestContext));

                _softAssertions.Add("Test Add Members", TestStatus.Success, AddMembersDL(TestContext));
                _softAssertions.Add("Test Verify Add Members", TestStatus.Success, VerifyMembersDL(TestContext, false));

                _softAssertions.Add("Test Add Administrator", TestStatus.Success, AddAdministratorDL(TestContext));
                _softAssertions.Add("Test Add Administrator", TestStatus.Success, VerifyMembersDL(TestContext, false));




                _softAssertions.Add("Test Add Email Address", TestStatus.Success,
                    AddAdditionalEmailAddress(TestContext));
                _softAssertions.Add("Test Verify Add Email Address", TestStatus.Success,
                    VerifyAdditionalEmailAddress(TestContext));


                _softAssertions.Add("Test Add SendOnBehalf Users", TestStatus.Success,
                    AddSendOnBehalfUsers(TestContext));
                _softAssertions.Add("Test Verify Add SendOnBehalf Users", TestStatus.Success,
                    VerifyAddSendOnBehalfUsers(TestContext));


                _softAssertions.Add("Test Add SendAsPermissions Users", TestStatus.Success,
                    AddSendAsPermissions(TestContext));
                _softAssertions.Add("Test Verify Add SendAsPermissions Users", TestStatus.Success,
                    VerifySendAsPermissions(TestContext));

                _softAssertions.Add("Test Add Accepted Senders Users", TestStatus.Success,
                    AddAcceptedSenders(TestContext));
                _softAssertions.Add("Test Add Accepted Senders Users", TestStatus.Success,
                    VerifyAcceptedSenders(TestContext));

                _softAssertions.Add("Test Add Rejected Users", TestStatus.Success,
                    AddRejectedSenders(TestContext));
                _softAssertions.Add("Test Verify Add Rejected Users", TestStatus.Success, 
                    VerifyRejectedSenders(TestContext));

               


                ClickDlBreakCrumb();

            }


        }
    }
}
