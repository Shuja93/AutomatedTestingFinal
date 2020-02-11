using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.TestTracker;

namespace HC10Test
{
    [TestClass]
    public class TestClassPublicFolder : BasePublicFolder
    {

        public TestContext TestContext { get; set; }
        private SoftAssertions _softAssertions;

        [ClassInitialize]
        public static void ClassSetup(TestContext TestContext)
        {

            OpenBrowser(BrowserType.Chrome);
            DriverContext.Browser.GoToUrl("https://hostingcontrollerdemo.com/");
            DriverContext.Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
            DriverContext.Driver.Manage().Window.Maximize();
            LoginPage login = new LoginPage();
            login.Login();
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

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", @"C:\Users\Shuja\source\repos\AutomationFramework\AutomatedTesting\HC10Test\Modules\Exchange\PublicFolder\Data\PublicFolder_Create.csv", "PublicFolder_Create#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        [TestCategory("Public Folder")] 
        public void PublicFolderCreation()
        {
            NavigateToPublicFolderPage(TestContext);
            _softAssertions.Add("Test Create Mailbox", TestStatus.Success, CreatePublicFolder(TestContext));
        }

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
            @"C:\Users\Shuja\source\repos\AutomationFramework\AutomatedTesting\HC10Test\Modules\Exchange\MailContacts\Data\MailContact_Dashboard.csv", "MailContact_Dashboard#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        [TestCategory("Public Folder")]
        public void PublicFolderUpdateDashboard()
        {
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
