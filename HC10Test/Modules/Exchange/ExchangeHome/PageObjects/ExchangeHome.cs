using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Helpers;
using HC10AutomationFramework.Extensions;

namespace HC10Test.PageObjects
{
    class ExchangeHome : BasePage
    {

        private IWebElement dropdownValue =>
            DriverContext.Driver.FindElement(By.XPath("//*[@id='RecsPerPage']/div[2]/select"));

        protected IWebElement searchBarOrganization =>
            DriverContext.Driver.FindElement(By.XPath("//*[@id='SystemName']"));

        private IWebElement btnMailbox =>
            ByXPath("//tr[1]//button[contains(@onclick, 'ExchangeOrganization.MailBoxes')]");

        private IWebElement btnDistributionLists =>
            ByXPath("//tr[1]//button[contains(@onclick, 'ExchangeOrganization.DistributionList')]");

        private IWebElement btnToggle => ByXPath("//tr[1]//button[@data-toggle='dropdown']");

        private IWebElement btnMailContact =>
            ByXPath("//tr[1]//li//a[contains(@onclick,'ExchangeOrganization.MailContacts')]");

        private IWebElement btnResourceMailbox =>
            ByXPath("//tr[1]//li//a[contains(@onclick,'ExchangeOrganization.ResourceMailbox')]");

        private IWebElement btnPublicFolderElem =>
            ByXPath("//tr[1]//li//a[contains(@onclick,'ExchangeOrganization.PublicFolders')]");





        public void SearchOrganizationName(string objectName)
        {
            SeleniumHelperMethods.ObjectSearchBar(DriverContext.Driver, searchBarOrganization, btnSearch,
                headerProgressElem, headerProgressElemBy, objectName);
        }

        public ExgOrgMailboxes MailboxesHome()
        {
            //btnMailbox.ClickWithWait("header"); 
            btnMailbox.Click();
            return new ExgOrgMailboxes();
        }

        public ExgOrgDL DistributionListsHome()
        {
            //btnMailbox.ClickWithWait("header"); 
            btnDistributionLists.Click();
            return new ExgOrgDL();
        }

        public ExgOrgMailContacts MailContactHome()
        {
            btnToggle.Click();
            btnMailContact.Click();
            return new ExgOrgMailContacts();
        }

        public ExgOrgResourceMailboxes ResourceMailboxHome()
        {
            btnToggle.Click();
            btnResourceMailbox.Click();
            return new ExgOrgResourceMailboxes();
        }

        public ExgOrgPublicFolders PublicFoldersHome()
        {
            btnToggle.Click();
            btnPublicFolderElem.Click();
            return new ExgOrgPublicFolders();



        }
    }


}
