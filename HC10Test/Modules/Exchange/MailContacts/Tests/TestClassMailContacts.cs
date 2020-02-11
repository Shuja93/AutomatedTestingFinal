using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HC10Test.PageObjects;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.TestTracker;

namespace HC10Test
{
    [TestClass]
    public class TestClassMailContacts : BaseMailContact
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

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", @"C:\Users\Shuja\source\repos\AutomationFramework\AutomatedTesting\HC10Test\Modules\Exchange\MailContacts\Data\MailContact_Create.csv", "MailContact_Create#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        [TestCategory("Mail Contact")] 
        public void MailContactCreation()
        {
            NavigateToMailContact(TestContext);
            _softAssertions.Add("Test Create Mailbox", TestStatus.Success, CreateMailContact(TestContext));


            if (TestTracker.mailContactStatus[Convert.ToString(TestContext.DataRow["InternalEmailAddress"])] == TestStatus.Success)
            {
                NavigateToMailContactDashboard(TestContext);
                _softAssertions.Add("Test Verify Mailbox General Properties", TestStatus.Success, VerifyMailContactGeneralProfile(TestContext, true));
                _softAssertions.Add("Test Verify External Email", TestStatus.Success, VerifyExternalEmail(TestContext));
                _softAssertions.Add("Test Verify External Email", TestStatus.Success, VerifyInternalEmail(TestContext));
                ClickMailContactrumb();
                //_softAssertions.Add("Test Verify Mailbox Advance Properties", TestStatus.Success, VerifyMailBoxAdvanceProperties(TestContext, true));
            }
        }

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
            @"C:\Users\Shuja\source\repos\AutomationFramework\AutomatedTesting\HC10Test\Modules\Exchange\MailContacts\Data\MailContact_Dashboard.csv", "MailContact_Dashboard#csv", DataAccessMethod.Sequential)]
        [TestMethod]
        [TestCategory("Exchange")]
        [TestCategory("Mail Contact")]
        public void MailContactUpdate_Dashboard()
        {
            if (TestTracker.mailContactStatus[Convert.ToString(TestContext.DataRow["InternalEmailAddress"])] == TestStatus.Success)
            {
                NavigateToMailContact(TestContext);
                NavigateToMailContactDashboard(TestContext);
                _softAssertions.Add("Test Add Email Address", TestStatus.Success,
                    AddAdditionalEmailAddress(TestContext));
                _softAssertions.Add("Test Add AcceptedSenders Users", TestStatus.Success,
                    AddAcceptedSenders(TestContext));
                _softAssertions.Add("Test Add Rejected Users", TestStatus.Success, AddRejectedSenders(TestContext));

                ClickMailContactrumb();

            }


        }
    }
}
