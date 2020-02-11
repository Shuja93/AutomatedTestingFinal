using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Helpers;
using HC10AutomationFramework.Logs;

namespace HC10AutomationFramework.Extensions
{
    public static class ElementExtensions
    {
        //revisit
        public static void EnterText(this IWebElement element, string text, string elementName)
        {

            element.Clear();
            element.SendKeys(text);
        }
        //revisit
        public static bool IsDisplayed(this IWebElement element, string elementName)
        {
            bool result;
            try
            {
                result = element.Displayed;
            }
            catch (Exception)
            {
                result = false;
            }

            return result;
        }

        public static void ClickOnIt(this IWebElement element, string elementName)
        {
            element.Click();
        }

        //revisit
        public static void SelectByText(this IWebElement element, string text, string elementName)
        {
            SelectElement oSelect = new SelectElement(element);
            oSelect.SelectByText(text);
        }

        //revisit
        public static void ClickWithWait(this IWebElement element,string waitMethod)
        {
            DriverContext.Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            element.Click();
            
            switch (waitMethod) 
            {
                case "header":
                    HeaderWait();
                    break;
                case "spinner":
                    SpinnerWait();
                    break;
                    
            }

            DriverContext.Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);

        }

        //revisit
        static void HeaderWait()
        {
            try
            {
                var wait = new WebDriverWait(DriverContext.Driver, TimeSpan.FromSeconds(2));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("headerProgressIndi")));
                while (true)
                {
                    if (DriverContext.Driver.FindElement(By.Id("headerProgressElem")).GetCssValue("display") == "none")
                    {
                        break;
                    }
                }
            }
            catch (Exception )
            {
                
            }
            
        }

        //revisit
        static void SpinnerWait() 
        {
            try
            {
                while (true)
                {
                    DriverContext.Driver.FindElement(By.CssSelector("spinnerbg"));
                }
            }
            catch (Exception )
            {

            }
        }

    }
}
