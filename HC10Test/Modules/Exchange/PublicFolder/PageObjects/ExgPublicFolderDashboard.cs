using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;
using OpenQA.Selenium;

namespace HC10Test.PageObjects
{
     class ExgPublicFolderDashboard : BasePage
    {

        //private IWebElement lnkEmailAddressElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#EditEmailAddresses']"));
        private IWebElement lnkAdvanceProperties => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SAdvancedProperties123456']"));
        private IWebElement lnkSetPermissions => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SetPermissions']"));

        private IWebElement btnSetPermissionsToAllUsersELem =>
            ByXPath("//button[contains(@onclick , 'PublicFolder.EditPermissionsForAll')]");


        private readonly DashboardEmailAddress pageEmailAddress;
        private readonly DashboardAccepetedSenders pageAcceptedSenders;
        private readonly DashboardRejectedSenders pageRejectedSenders;
        private readonly DashboardGeneralProfile pageGeneralProfile;
       // private readonly DashboardAdvanceProperties pageAdvanceProperties;
        private readonly DashboardForwarding pageForwarding;

        public ExgPublicFolderDashboard()
        {
            pageEmailAddress = new DashboardEmailAddress();
            pageAcceptedSenders = new DashboardAccepetedSenders();
            pageRejectedSenders = new DashboardRejectedSenders();
            pageGeneralProfile = new DashboardGeneralProfile();
            //pageAdvanceProperties = new DashboardAdvanceProperties();
            pageForwarding = new DashboardForwarding();
        }

        public string VerifyGeneralProperties(string firstname, string lastName, string displayName, string country, string state, string
            officeLocation, string address, string city, string zipCode, string jobTitle, string company, string department, string managedBy, string businessPhone, string
            fax, string homePhone, string mobilePhone, string pager, string notes) => pageGeneralProfile.VerifyGeneralProperties(firstname, lastName, displayName, country, state,
            officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
            fax, homePhone, mobilePhone, pager, notes);

        public string SetGeneralProperties(string firstname, string lastName, string displayName, string country,
            string state, string officeLocation, string address, string city, string zipCode, string jobTitle,
            string company, string department, string managedBy, string businessPhone, string fax, string homePhone,
            string mobilePhone, string pager, string notes) => pageGeneralProfile.SetGeneralProperties(firstname,
            lastName, displayName, country, state,
            officeLocation, address, city, zipCode, jobTitle, company, department, managedBy, businessPhone,
            fax, homePhone, mobilePhone, pager, notes);

        public string SetAdditionalEmailAddress(string additionalEmail) => pageEmailAddress.SetAdditionalEmailAddress(additionalEmail);
        public string VerifyAdditionalEmailAddress(string additionalEmail) => pageEmailAddress.VerifyAdditionalEmailAddress(additionalEmail);
        public string SetAcceptedSenders(string userList) => pageAcceptedSenders.SetAcceptedSenders(userList);
        public string VerifyAcceptedSenders(string userList) => pageAcceptedSenders.VerifyAcceptedSenders(userList);
        public string SetRejectedSenders(string userList) => pageRejectedSenders.SetRejectedSenders(userList);
        public string VerifyRejectedSenders(string userList) => pageRejectedSenders.VerifyRejectedSenders(userList);
        public string SetForwarding(string user, string ou, string exchangeObject, IWebElement forwardingButton) => pageForwarding.SetForwarding(user, ou, exchangeObject, forwardingButton);
        public string VerifyForwarding(string user) => pageForwarding.VerifyForwarding(user);













        public void OpenAdvancedProperties()
        {
            lnkAdvanceProperties.ClickWithWait("spinner");
        }

        public void OpenEmailAddress()
        {
            lnkEmailAddressElem.ClickWithWait("spinner");
        }

        public void OpenSetPermissions()
        {
            lnkSetPermissions.ClickWithWait("spinner");
        }


    }
}
