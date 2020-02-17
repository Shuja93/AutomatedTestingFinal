using System;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;
using OpenQA.Selenium;

namespace HC10Test.PageObjects
{
     class ExgMailContactDashboard : BasePage
    {
        private readonly DashboardEmailAddress pageEmailAddress;
        private readonly DashboardAccepetedSenders pageAcceptedSenders;
        private readonly DashboardRejectedSenders pageRejectedSenders;
        private readonly DashboardGeneralProfile pageGeneralProfile;
       private readonly ExgMailContactAdvanceProperties pageAdvanceProperties;

        public ExgMailContactDashboard()
        {
            pageEmailAddress = new DashboardEmailAddress();
            pageAcceptedSenders = new DashboardAccepetedSenders();
            pageRejectedSenders = new DashboardRejectedSenders();
            pageGeneralProfile = new DashboardGeneralProfile();
            pageAdvanceProperties = new ExgMailContactAdvanceProperties();
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

        public string SetAdvanceProperties(string displayName, string externalEmailAddress, bool isHiddenFromAddressList, string maximumRecipients, string maximumReceiveSize) => pageAdvanceProperties.SetAdvanceProperties(displayName, externalEmailAddress, isHiddenFromAddressList, maximumRecipients, maximumReceiveSize);


        public string SetAdditionalEmailAddress(string additionalEmail) => pageEmailAddress.SetAdditionalEmailAddress(additionalEmail);
        public string VerifyAdditionalEmailAddress(string additionalEmail) => pageEmailAddress.VerifyAdditionalEmailAddress(additionalEmail);
        public string SetAcceptedSenders(string userList) => pageAcceptedSenders.SetAcceptedSenders(userList);
        public string VerifyAcceptedSenders(string userList) => pageAcceptedSenders.VerifyAcceptedSenders(userList);
        public string SetRejectedSenders(string userList) => pageRejectedSenders.SetRejectedSenders(userList);
        public string VerifyRejectedSenders(string userList) => pageRejectedSenders.VerifyRejectedSenders(userList);

        public string VerifyExternalEmailAddress(string expectedExternalEmail)
        {
            try
            {
                string actualExternalEmail = DriverContext.Driver.FindElement(By.XPath("//td[4]")).Text;
                if (actualExternalEmail == expectedExternalEmail)
                {
                    return TestStatus.Success;
                }
                else
                {
                    return actualExternalEmail;
                }

            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        public string VerifyInternalEmailAddress(string externalEmail, string expectedInternalEmail)
        {
            if (string.IsNullOrWhiteSpace(expectedInternalEmail))
            {
                expectedInternalEmail = externalEmail;
            }
            try
            {
                string actualInternalEmail = DriverContext.Driver.FindElement(By.XPath("//td[5]")).Text;
                if (actualInternalEmail == expectedInternalEmail)
                {
                    return TestStatus.Success;
                }
                else
                {
                    return actualInternalEmail;
                }

            }
            catch (Exception e)
            {
                return e.Message;
            }
        }






        private IWebElement lnkEmailAddressElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#EditEmailAddresses']"));
        private IWebElement lnkAdvanceProperties => DriverContext.Driver.FindElement(By.CssSelector("[href*='#EditAdv']"));




        public void OpenAdvancedProperties()
        {
            lnkAdvanceProperties.ClickWithWait("spinner");
        }

        public void OpenEmailAddress()
        {
            lnkEmailAddressElem.ClickWithWait("spinner");
        }


    }
}
