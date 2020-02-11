using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using System.Drawing.Text;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Helpers;

namespace HC10Test.PageObjects
{
     class ExgCreatePublicFolder : BasePage
     {
         private IWebElement txtPublicFolderNameElem =>
             ByXPath("//input[@id = 'PFName' and @data-val-required = 'Please enter a valid public folder name.']");

         private IWebElement dropdownPublicFolderTyoeElem => ByXPath("//*[@id='PublicFolderType']");
         private IWebElement ckbxPFMailEnableElem => ByXPath("//*[@id='IsMailEnabled']");
         private IWebElement txtEmailAddressElem => ByXPath("//*[@id='DisplayName']");
         private IWebElement txtSetQuotaElem => ByXPath("//*[@id='SetQuota']/input[@type = 'text']");
         private IWebElement ckbxSetQuotaElem => ByXPath("//*[@id='SetQuota']//input[@type = 'checkbox']");
         private IWebElement btnCreatePublicFolder => ByXPath("//button[@type='submit']");
         private IWebElement dropdownPFMailboxElem => ByXPath("//*[@id='PFMailbox']");



        public string CreatePublicFolder(string publicFolderName, string publicFolderType, bool isMailEnable,
            string email, string publicFolderMailbox, string publicFolderSize, bool isCr)
        {

            try
            {
                txtPublicFolderNameElem.SendKeys(publicFolderName);

                
                MailAddress addr = new MailAddress(email);
                string userName = addr.User;
                string mailDomain = addr.Host;
                txtEmailAddressElem.SendKeys(userName);
                

                if (publicFolderType != "Mail")
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownPublicFolderTyoeElem, publicFolderType);
                }

                if (isMailEnable == true)
                {
                    if (ckbxPFMailEnableElem.Selected)
                    {
                        
                    }
                    else
                    {
                        ckbxPFMailEnableElem.Click();
                    }
                }

                if (!string.IsNullOrEmpty(publicFolderMailbox))
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownPFMailboxElem, publicFolderMailbox);
                }

                if (isCr == true)
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, publicFolderSize);
                }

                else
                {
                    if (ckbxSetQuotaElem.Selected)
                    {
                        ckbxPFMailEnableElem.Click();
                    }
                    txtSetQuotaElem.Clear();
                    txtSetQuotaElem.SendKeys(publicFolderSize);
                }

                Thread.Sleep(2000);
                    btnCreatePublicFolder.Click();
                    
                
                
                return  GetPrompt(headerProgressElem, headerProgressElemBy, MessageContainer.DialogeContainer);
                
            }

            

            catch (Exception e)
            {
                return e.Message.Trim();
            }

            
        }

    }
}
