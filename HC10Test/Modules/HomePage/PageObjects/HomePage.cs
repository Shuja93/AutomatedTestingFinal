using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Extensions;

namespace HC10Test.PageObjects
{
    class HomePage : BasePage
    {
        //Side Menu

        //Provisioning Tab
        private IWebElement lnkProvisioning => DriverContext.Driver.FindElement(By.XPath("//*[@id='side-menu']/li[contains(.,'Provisioning')]"));


        //Sub Menu
        private IWebElement lnkExchange => DriverContext.Driver.FindElement(By.LinkText("Exchange"));

       




       

        //Exchange Tab
        

        public void ClickProvisioning()
        {
            //lnkProvisioning.Click();
            lnkProvisioning.Click();
        }

        public ExchangeHome ClickExchangeHome()
        {
            //lnkExchange.ClickWithWait("header");
            lnkExchange.Click();
            return new ExchangeHome();
        }
    }


}
