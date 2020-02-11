using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Helpers;
using System.Net.Mail;
using HC10AutomationFramework.Extensions;
using HC10Test.PageObject;

namespace HC10Test.PageObjects  
{
     class ExgOrgDL : BasePage
    {
        private IWebElement btnCreateDistibutionList => DriverContext.Driver.FindElement(By.XPath("//button[contains(@onclick, 'DistributionList.AddDistributionList')]"));
        private IWebElement btnDLDashboardElem => DriverContext.Driver.FindElement(By.XPath("//tr[1]//button[contains(@onclick, 'DistributionList.ShowDashboard(this);')]"));
        private IWebElement searchBarMailboxElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='DisplayName']"));

       

        public ExgCreateDistributionList OpenCreateDLPage()
        {
            btnCreateDistibutionList.Click();
            return new ExgCreateDistributionList();
        }

        public void SearchName(string email, string displayName)
        {
            string searchString;

            if (displayName == "")
            {
                MailAddress addr = new MailAddress(email);
                searchString = addr.User;
            }
            else
            {
                searchString = displayName;
            }


            SeleniumHelperMethods.ObjectSearchBar(DriverContext.Driver, searchBarMailboxElem, btnSearch, headerProgressElem, headerProgressElemBy, searchString);
        }

        public ExgDLDashboard OpenDLDashboard()
        {
            btnDLDashboardElem.ClickWithWait("header");
            return new ExgDLDashboard();
        }
        public void SearchDL(string email, string displayName)
        {
            string searchString;

            if (displayName == "")
            {
                MailAddress addr = new MailAddress(email);
                searchString = addr.User;
            }
            else
            {
                searchString = displayName;
            }


            SeleniumHelperMethods.ObjectSearchBar(DriverContext.Driver, searchBarMailboxElem, btnSearch, headerProgressElem, headerProgressElemBy, searchString);
        }
    }
}
