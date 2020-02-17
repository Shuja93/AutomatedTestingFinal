using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Enum;
using OpenQA.Selenium;

namespace HC10Test.PageObjects
{
    public class ExgMailContactAdvanceProperties : BasePage
    {
        private IWebElement txtDisplayNameElem => ByXPath("//*[@id='DisplayName']");
        private IWebElement txtExternalEmailAddressElem => ByXPath("//*[@id='ExternalEmailAddress']");
        private IWebElement ckbxHideFromAddressListElem => ByXPath("//*[@id='HiddenFromAddressLists']");
        private IWebElement txtMaximumRecepientsElem => ByXPath("//*[@id='MaximumRecipients']//input[@type = 'text']");
        private IWebElement txtMaximumReceiveSizeElem => ByXPath("//*[@id='MaxReceiveSize']//input[@type = 'text']");

        private IWebElement btnSaveMailContactAdvanceProperties =>
            ByXPath("//*[@id='advanceProperties']//button[@type = 'submit']");

        public string SetAdvanceProperties(string displayName, string externalEmailAddress, bool isHiddenFromAddressList, string maximumRecipients,string maximumReceiveSize)
        {
            try
            {
                if (string.IsNullOrEmpty(displayName))
                {
                    txtDisplayNameElem.Clear();
                    txtDisplayNameElem.SendKeys(displayName);
                }

                if (string.IsNullOrEmpty(externalEmailAddress))
                {
                    txtExternalEmailAddressElem.Clear();
                    txtExternalEmailAddressElem.SendKeys(externalEmailAddress);
                }
                

                SetCheckBox(ckbxHideFromAddressListElem,isHiddenFromAddressList);

                if (string.IsNullOrEmpty(maximumRecipients))
                {
                    txtMaximumRecepientsElem.Clear();
                    txtMaximumRecepientsElem.SendKeys(maximumRecipients);
                }

                if (string.IsNullOrEmpty(maximumReceiveSize))
                {
                    txtMaximumReceiveSizeElem.Clear();
                    txtMaximumReceiveSizeElem.SendKeys(maximumReceiveSize);
                }

                btnSaveMailContactAdvanceProperties.Click();

                return GetPrompt(headerProgressElem, headerProgressElemBy, MessageContainer.ToastContainer);
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

    }
}
