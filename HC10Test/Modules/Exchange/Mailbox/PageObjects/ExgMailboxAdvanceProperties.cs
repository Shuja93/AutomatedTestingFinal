using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Helpers;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace HC10Test.PageObjects
{
    public class ExgMailboxAdvanceProperties : BasePage
    {
        private string accumulatedQuota => "I'll choose my own offering";
        private IWebElement ckbxHideFromAddressBookElem => ByXPath("//*[@id='HideFromAddressBook']");
        private IWebElement ckbxImapElem => ByXPath("//*[@id='IMAP']");
        private IWebElement ckbxPopElem => ByXPath("//*[@id='POP']");
        private IWebElement ckbxOwaElem => ByXPath("//*[@id='OWA']");
        private IWebElement ckbxMapiElem => ByXPath("//*[@id='MAPI']");

        private IWebElement btnUpdateAdvProp =>
            ByXPath("//*[@id='SAdvancedProperties123456']//button[@type = 'submit']");

        public string VerifyAdvancedProperties(string mailboxSize, bool isCR)
        {
            try
            {
                expectedProperties.Add("MailboxSize", mailboxSize);
                expectedProperties.Add("IncomingSize", mailboxSize);
                expectedProperties.Add("OutgoingSize", mailboxSize);
                expectedProperties.Add("ProhibitSendAt", mailboxSize);
                expectedProperties.Add("IssueWarningAt", mailboxSize);
                expectedProperties.Add("Quota", isCR == false ? "Accumulated" : mailboxSize);


                actualProperties.Add("MailboxSize", Convert.ToString(txtMailboxSizeElem.GetAttribute("value")));
                actualProperties.Add("IncomingSize", Convert.ToString(txtIncomingSize.GetAttribute("value")));
                actualProperties.Add("OutgoingSize", Convert.ToString(txtOutgoingSize.GetAttribute("value")));
                actualProperties.Add("ProhibitSendAt", Convert.ToString(txtProhibitSendAtElem.GetAttribute("value")));
                actualProperties.Add("IssueWarningAt", Convert.ToString(txtWarnAtElem.GetAttribute("value")));
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

        public string SetAdvanceProperties(bool isCR, string mailboxSize,bool isHiddenFromAddressBook, bool isImapEnabled, bool isPopEnabled, bool isOwaEnabled, bool isMapiEnabled)
        {
            try
            {
                SetCheckBox(ckbxHideFromAddressBookElem,isHiddenFromAddressBook);

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
                    SetMailboxSize(mailboxSize);
                }

                btnUpdateAdvProp.Click();
                return GetPrompt(headerProgressElem, headerProgressElemBy, MessageContainer.ToastContainer);


            }
            catch (Exception e)
            {
                return e.Message.Trim();
            }
        }
    }
}
