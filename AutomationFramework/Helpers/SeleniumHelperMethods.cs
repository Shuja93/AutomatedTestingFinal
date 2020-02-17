using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Linq;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Logs;
namespace HC10AutomationFramework.Helpers
{
    public class SeleniumHelperMethods
    {
        public static bool SelectDropDownValue(IWebElement dropDownElement, string value)
        {
            try
            {
                var selectElement = new SelectElement(dropDownElement);
                selectElement.SelectByText(value);
                return true;
            }
            catch (Exception e)
            {
                return false;
                LogClass.AppendLogs(e);
            }
            
        }

        

       

        public static void LoadWait(IWebDriver driver,IWebElement loadWaitElement, string loadWaitString) 
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id(loadWaitString)));
                while (true)
                {
                    if (loadWaitElement.GetCssValue("display") == "none") { break; }
                }
            }
            catch (TimeoutException)
            {
                
            }
            
        }

        

        public static bool WaitExpectedConditionsClickable(IWebDriver driver, IWebElement element) 
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(element));
                return true;
            }
            catch (TimeoutException )
            {
                return false;
            }
            
            
        }

        public static void ObjectSearchBar(IWebDriver driver, IWebElement searchBar, IWebElement btnSearch, IWebElement loadwaitElement, string loadwaitString,  string objectName) 
        {
            searchBar.SendKeys(objectName);
            btnSearch.Click();
            //var wait = new WebDriverWait(DriverContext.Driver, new TimeSpan(0, 0, 30));
            //var element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='headerProgressIndi']")));
            SeleniumHelperMethods.LoadWait(driver, loadwaitElement, loadwaitString);
        }


    }
}
