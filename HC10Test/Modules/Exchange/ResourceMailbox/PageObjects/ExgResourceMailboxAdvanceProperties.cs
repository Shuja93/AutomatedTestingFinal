using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using System.Threading;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Helpers;
using OpenQA.Selenium.Support.UI;

namespace HC10Test.PageObjects
{
    class ExgResourceMailboxAdvanceProperties : BasePage
    {
        private string accumulatedQuota => "I'll choose my own offering";
        private IWebElement ckbxHideFromAddressBookElem => ByXPath("//*[@id='HideFromAddressBook']");
        private IWebElement ckbxImapElem => ByXPath("//*[@id='IMAP']");
        private IWebElement ckbxPopElem => ByXPath("//*[@id='POP']");
        private IWebElement ckbxOwaElem => ByXPath("//*[@id='OWA']");
        private IWebElement ckbxMapiElem => ByXPath("//*[@id='MAPI']");

        private IWebElement btnUpdateAdvProp =>
            ByXPath("//*[@id='SAdvancedProperties123456']//button[@type = 'submit']");


        private IWebElement ckbxUnlimitedMailboxSize => ByXPath("//*[@id='MailboxSize']//input[@type = 'checkbox']");
        private IWebElement txtMailboxSize => ByXPath("//*[@id='MailboxSize']//input[@type = 'text']");

        private IWebElement ckbxIncomingMessageSize =>
            ByXPath("//*[@id='MaxIncomingMsgSize']//input[@type = 'checkbox']");

        private IWebElement txtIncomingMessageSize => ByXPath("//*[@id='MaxIncomingMsgSize']//input[@type = 'text']");
        private IWebElement ckbxOutGoingSize => ByXPath("//*[@id='MaxOutgoingMsgSize']//input[@type = 'checkbox']");
        private IWebElement txtOutgoingMessageSize => ByXPath("//*[@id='MaxOutgoingMsgSize']//input[@type = 'text']");




        public string SetAdvanceProperties(bool isCR, string mailboxSize, bool isHiddenFromAddressBook, bool isImapEnabled, bool isPopEnabled, bool isOwaEnabled, bool isMapiEnabled)
        {
            try
            {
                SetCheckBox(ckbxHideFromAddressBookElem, isHiddenFromAddressBook);

                if (isCR == true)
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, mailboxSize);
                }

                else
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, accumulatedQuota);

                    SetCheckBox(ckbxImapElem, isImapEnabled);
                    SetCheckBox(ckbxPopElem, isPopEnabled);
                    SetCheckBox(ckbxOwaElem, isOwaEnabled);
                    SetCheckBox(ckbxMapiElem, isMapiEnabled);
                    SetResourceMailboxSize(mailboxSize);

                }

                btnUpdateAdvProp.Click();
                return GetPrompt(headerProgressElem, headerProgressElemBy, MessageContainer.ToastContainer);


            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        private void SetResourceMailboxSize(string mailboxSize)
        {
            SetCheckBox(ckbxUnlimitedMailboxSize,false);
            SetCheckBox(ckbxIncomingMessageSize,false);
            SetCheckBox(ckbxOutGoingSize,false);

            txtMailboxSize.Clear();
            txtMailboxSize.SendKeys(mailboxSize);

            txtIncomingMessageSize.Clear();
            txtIncomingMessageSize.SendKeys(mailboxSize);

            txtOutgoingMessageSize.Clear();
            txtOutgoingMessageSize.SendKeys(mailboxSize);
        }
        public string VerifyAdvancedProperties(string mailboxSize, bool isCR)
        {
            try
            {
                expectedProperties.Add("MailboxSize", mailboxSize);
                expectedProperties.Add("IncomingSize", mailboxSize);
                expectedProperties.Add("OutgoingSize", mailboxSize);
                expectedProperties.Add("Quota", isCR == false ? "Accumulated" : mailboxSize);


                actualProperties.Add("MailboxSize", Convert.ToString(txtMailboxSizeElem.GetAttribute("value")));
                actualProperties.Add("IncomingSize", Convert.ToString(txtIncomingSize.GetAttribute("value")));
                actualProperties.Add("OutgoingSize", Convert.ToString(txtOutgoingSize.GetAttribute("value")));

                
                var selectElement = new SelectElement(dropdownCRElem);
                string quotaStatus = selectElement.SelectedOption.Text;
                actualProperties.Add("Quota",
                    quotaStatus == "I'll choose my own offering" ? "Accumulated" : quotaStatus);


                return CompareLists(expectedProperties, actualProperties);

            }

            catch (Exception ex)
            {
                return ex.Message;
            }
        }



    }
}
