using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using System.Net.Mail;
using System.Threading;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Helpers;
using HC10AutomationFramework.Resources;

namespace HC10Test.PageObject
{
     class ExgCreateDistributionList : BasePage
    {
        private IWebElement radioDistributionElem => ByXPath("//input[@value= 'distribution']");
        private IWebElement radioSecurityElem => ByXPath("//input[@value= 'security']");
        private IWebElement radioDynamicDLElem => ByXPath("//input[@value= 'dynamicdistribution']");
        private IWebElement radioNewListElem => ByXPath("//*[@id='NewList']");
        private IWebElement radioExistingDLElem => ByXPath("//*[@id='ExistingList']");
        private IWebElement txtDLDisplayNameElem => ByXPath("//*[@id='newuser']//*[@id='DisplayName']");

        private IWebElement txtDLEmailAddressElem =>
            ByXPath("//*[@id='EmailAddress' and @class = 'form-control mandatory']");

        private IWebElement btnChooseExistingElem =>
            ByXPath("//button[contains(@onclick, 'DistributionList.GetExistingDistributionList')]");

        private IWebElement ckbxIsAllSenderAuthenticatedElem => ByXPath("//*[@id='IsAllSenderAuthenticated']");
        private IWebElement btnAddAdministratorElem => ByXPath("//*[@id='btnAdmin']");
        private IWebElement btnAddMembersElem => ByXPath("//*[@id='btnMember']");
        private IWebElement btnCreateLDElemElem => ByXPath("//button[@type = 'submit']");

        private IWebElement ckbxUnlimitedMessageSize =>
            ByXPath("//*[@id='InComingMessageSizeLimit']//input[@type = 'checkbox']");
        private IWebElement txtUnlimitedMessageSize =>
            ByXPath("//*[@id='InComingMessageSizeLimit']//input[@type = 'text']");

        public string CreateDl(string groupType, bool isSubOU, bool isNewGroup, string adGroupName, string email,
            string incomingMessageSize, bool isCR, string displayName, string adminUsers,
            string memberUsers, bool allSendersAuth)
        {
            try
            {
                SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCreateLDElemElem);

                switch (groupType)
                {
                    case "Distribution":
                        radioDistributionElem.Click();
                        break;
                    case "Security":
                        radioSecurityElem.Click();
                        break;
                    case "Dynamic":
                        radioDynamicDLElem.Click();
                        break;
                }
                if (isSubOU == true)
                {
                    if (!SelectSubOU(DriverContext.Driver, lnkSubOUElem, subOUWaitElem,
                         subOUWaitBy))
                    {
                        return ErrorDescriptions.ErrorSubOuNotFound;
                    }
                }

                if (isNewGroup)
                {
                    txtDLDisplayNameElem.SendKeys(displayName);
                }
                else
                {
                    radioExistingDLElem.Click();
                    btnChooseExistingElem.Click();
                    SelectExistingObject(DriverContext.Driver, adGroupName);
                }

                MailAddress addr = new MailAddress(email);
                string userName = addr.User;
                string mailDomain = addr.Host;
                txtDLEmailAddressElem.Clear();
                txtDLEmailAddressElem.SendKeys(userName);

                

                if (groupType == "Distribution" || groupType == "Security")
                {
                    AddDlMembers(memberUsers);
                    if (allSendersAuth)
                    {
                        ckbxIsAllSenderAuthenticatedElem.Click();
                    }
                }

                AddDlAdministrators(adminUsers);
                

                if (isCR == true)
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, incomingMessageSize);
                }
                else
                {
                    if (ckbxUnlimitedMessageSize.Selected == true)
                    { ckbxUnlimitedMessageSize.Click();}
                    txtUnlimitedMessageSize.Clear();
                    txtUnlimitedMessageSize.SendKeys(incomingMessageSize);
                }

                Thread.Sleep(2000);
                btnCreateLDElemElem.Click();
                //PageRefresh(DriverContext.Driver);

                return GetPrompt(headerProgressElem, headerProgressElemBy, MessageContainer.DialogeContainer);

            }
            catch (Exception e)
            {
                return e.Message.Trim();
            }
        }

        public void AddDlAdministrators(string userList)
        {
            btnAddAdministratorElem.Click();
            AddUsersinNewWindows(DriverContext.Driver, userList);
        }

        public void AddDlMembers(string userList)
        {
            btnAddMembersElem.Click();
            AddMembersInDl(DriverContext.Driver, userList);
        }




     }
}
