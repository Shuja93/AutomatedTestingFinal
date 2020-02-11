using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using System.Collections.Generic;

namespace HC10AutomationFramework.Base
{
    public class TestBase : SharedElementsPage
    {
       
        public static void OpenBrowser(BrowserType browserType)
        {
            switch (browserType)
            {
                case BrowserType.Edge:
                    DriverContext.Driver = new EdgeDriver();
                    DriverContext.Browser = new Browser(DriverContext.Driver);
                    break;

                case BrowserType.Firefox:
                    DriverContext.Driver = new FirefoxDriver();
                    DriverContext.Browser = new Browser(DriverContext.Driver);
                    break;

                case BrowserType.Chrome:
                    DriverContext.Driver = new ChromeDriver();
                    DriverContext.Browser = new Browser(DriverContext.Driver);
                    break;

                default:
                    DriverContext.Driver = new ChromeDriver();
                    DriverContext.Browser = new Browser(DriverContext.Driver);
                    break;
            }
        }


      
    }
}
