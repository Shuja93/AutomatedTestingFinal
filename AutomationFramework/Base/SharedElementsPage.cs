using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Helpers;
using OpenQA.Selenium;
using HC10AutomationFramework.Logs;
using System;
using System.Threading;

namespace HC10AutomationFramework.Base
{
    public class SharedElementsPage
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
        protected IWebElement colorBoxElem => ById("colorbox");
        protected string colorBoxElemBy = "colorbox";

        protected IWebElement lnkSubOUElem => ByXPath("//*[@id='OUPath']");
        protected IWebElement subOUWaitElem => ByXPath("//*[@id='btnSubOU']");
        protected string subOUWaitBy = "btnSubOU";
        protected IWebElement btnProfileArrowToggleElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='MainHeader']/nav/div[2]/ul/li[3]/a/span"));
        protected IWebElement btnLogoutElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='MainHeader']/nav/div[2]/ul/li[3]/ul/li[8]/a"));
        protected IWebElement btnSearch => DriverContext.Driver.FindElement(By.XPath("//*[@id='SearchSubmit']"));
       
        protected IWebElement txtFirstNameElem => ById("GeneralProfile_FirstName");
        protected IWebElement txtLastNameElem => ById("GeneralProfile_LastName");
        protected IWebElement txtDisplayNameElem => ById("GeneralProfile_DisplayName");
        protected IWebElement dropdownCountryElem => ByXPath("//*[@id='select2-Country-container']");
        protected IWebElement txtCountryElem => ByXPath("//*[contains(@class, 'select2-search__field')]");
        protected IWebElement dropdownStateElem => ByXPath("//*[@id='GeneralProfile_ExistingState']");
        protected IWebElement btnVerifyDisableElem => ByXPath("/html/body/div[5]/div/button[1]");
        protected IWebElement dialogueContainerElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='DialogStatusMessageContainer']/div"));
        //protected IWebElement dialogueContainerElem => DriverContext.Driver.FindElement(By.XPath("//*[@id='DialogStatusMessageContainer']/div/text()"));
        
        protected IWebElement toastContainer => ByXPath("//*[@id='toast-container']/div/div[2]");
        protected IWebElement lnkadvancedPropertiesElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SAdvancedProperties123456']"));
        protected IWebElement btnTabRefreshButtonElem => ByXPath("//i[@title = 'Reload Current Tab']");
        protected IWebElement dropdownExchangeObjects => ById("MailboxType");
        protected IWebElement dropdownOU => ById("OrganizationName");
        protected IWebElement btnCreateMailboxElem => ByXPath("//*[@id='SaveBtn']");
        protected IWebElement spinnerbgElem => DriverContext.Driver.FindElement(By.CssSelector("spinnerbg"));

        protected IWebElement lnkMailContact => DriverContext.Driver.FindElement(By.LinkText("Mail Contacts"));
        protected IWebElement lnkMailboxes => DriverContext.Driver.FindElement(By.LinkText("Mailboxes"));
        protected IWebElement lnkResourceMailboxes => DriverContext.Driver.FindElement(By.LinkText("Resource Mailboxes"));
        protected IWebElement lnkPublicFolder => DriverContext.Driver.FindElement(By.LinkText("Public Folders"));
        protected IWebElement lnkDl => DriverContext.Driver.FindElement(By.LinkText("Distribution Lists"));

        protected IWebElement txtMailboxSizeElem => ByXPath("//*[@id='MailboxSize']/input[contains(@type,'text')]");
        protected IWebElement ckbxMailboxSizeUnlimitedElem => ByXPath("//*[@id='MailboxSize']//input[@type = 'checkbox']");
        protected IWebElement txtProhibitSendAtElem => ByXPath("//*[@id='ProhibitSend']/input[contains(@type,'text')]");
        protected IWebElement txtWarnAtElem => ByXPath("//*[@id='WarnAt']//input[contains(@type,'text')]");
        protected IWebElement txtIncomingSize => ByXPath("//*[@id='MaxIncomingMsgSize']/input[contains(@type,'text')]");
        protected IWebElement txtOutgoingSize => ByXPath("//*[@id='MaxOutgoingMsgSize']/input[contains(@type,'text')]");
        protected IWebElement ckbxIncomingSizeUnlimitedElem => ByXPath("//*[@id='MaxIncomingMsgSize']//input[@type = 'checkbox']");
        protected IWebElement ckbxOutgoingSizeUnlimited => ByXPath("//*[@id='MaxOutgoingMsgSize']//input[@type = 'checkbox']");
        protected IWebElement dropdownCRElem => ById("CResourceId");
        protected IWebElement btnAddStateElem => ByXPath("//*[@id='existingSpan']/button");
        protected IWebElement txtNewStateElem => ByXPath("//*[@id='GeneralProfile_NewState']");
        protected IWebElement btnCloseDialogueBox =>
            ByXPath("//button[contains(text(), 'Cancel') and contains(@onclick, 'Layout.CloseModelDialog()')]");

        protected IWebElement lnkEmailAddressElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#PEmailAddress123456']"));
        protected IWebElement lnkMembershipElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#Membership']"));
        protected IWebElement lnkSendOnBehalfElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SendOnBehalf']"));
        protected IWebElement lnkFullAccessPermissionsElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#FullAccessPermission']"));
        protected IWebElement lnkSenAsPermissionsElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#SendAsPermission']"));
        protected IWebElement lnkAcceptedSendersElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#AcceptedSenders']"));
        protected IWebElement lnkRejectedSendersElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#RejectedSenders']"));
        protected IWebElement lnkForwardingElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#Forwarding']"));

    }
}
