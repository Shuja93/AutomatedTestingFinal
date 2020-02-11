using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Helpers;
using OpenQA.Selenium;
using HC10AutomationFramework.Logs;
using System;
using System.Threading;

namespace HC10AutomationFramework.Base
{
    public class ElementsBasePage
    {
        public IWebElement ById(string elemId)
        {
            return DriverContext.Driver.FindElement(By.Id(elemId));
        }

        public IWebElement ByXPath(string elemXPath)
        {
            return DriverContext.Driver.FindElement((By.XPath(elemXPath)));
        }

        protected IWebElement headerProgressElem => ById("headerProgressIndi");
        protected string headerProgressElemBy = "headerProgressIndi";
        protected IWebElement lnkSubOUElem => ByXPath("//*[@id='OUPath']");
        protected IWebElement subOUWaitElem => ByXPath("//*[@id='btnSubOU']");
        protected string subOUWaitBy = "btnSubOU";
        protected IWebElement btnProfileArrowToggleElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='MainHeader']/nav/div[2]/ul/li[3]/a/span"));
        protected IWebElement btnLogoutElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='MainHeader']/nav/div[2]/ul/li[3]/ul/li[8]/a"));
        protected IWebElement btnSearch => DriverContext.Driver.FindElement(By.XPath("//*[@id='SearchSubmit']"));
        protected IWebElement txtMailboxSizeElem => ByXPath("//*[@id='MailboxSize']/input[1]");
        protected IWebElement txtProhibitSendAtElem => ByXPath("//*[@id='ProhibitSend']/input[1]");
        protected IWebElement txtWarnAtElem => ByXPath("//*[@id='WarnAt']/input[1]");
        protected IWebElement txtIncomingSize => ByXPath("//*[@id='MaxIncomingMsgSize']/input[1]");
        protected IWebElement txtOutgoingSize => ByXPath("//*[@id='MaxOutgoingMsgSize']/input[1]");
        protected IWebElement ckbxIncomingSizeUnlimitedElem => ByXPath("//*[@id='MaxIncomingMsgSize']/label/input");
        protected IWebElement ckbxOutgoingSizeUnlimited => ByXPath("//*[@id='MaxOutgoingMsgSize']/label/input");
        protected IWebElement dropdownCRElem => ById("CResourceId");
        protected IWebElement txtFirstNameElem => ById("GeneralProfile_FirstName");
        protected IWebElement txtLastNameElem => ById("GeneralProfile_LastName");
        protected IWebElement txtDisplayNameElem => ById("GeneralProfile_DisplayName");
        protected IWebElement dropdownCountryElem => ByXPath("//*[@id='select2-Country-container']");
        protected IWebElement txtCountryElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='colorbox']/span/span/span[1]/input"));
        protected IWebElement dropdownStateElem => ByXPath("//*[@id='GeneralProfile_ExistingState']");
        protected IWebElement btnVerifyDisableElem => ByXPath("/html/body/div[5]/div/button[1]");
        protected IWebElement dialogueContainerElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='DialogStatusMessageContainer']/div"));
        protected IWebElement toastContainer => ByXPath("//*[@id='toast-container']/div/div[2]");
        protected IWebElement advancedProperties => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SAdvancedProperties123456']"));



    }
}
