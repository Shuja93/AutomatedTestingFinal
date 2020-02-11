using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Helpers;

namespace HC10Test.PageObjects
{
     class ExgCreateMailContact : BasePage
    {
        //private IWebElement txtContactNameElem => ByXPath("//*[@id='ContactName']");
        private IWebElement txtContactNameElem => ByXPath("//*[@id='ContactName' and @data-val-required='Please enter contact name.']");
        private IWebElement txtFirstNameElem => ByXPath("//*[@id='FirstName']");
        private IWebElement txtLastNameElem => ByXPath("//*[@id='LastName']");
        private IWebElement txtExternalEmailAddressElem => ByXPath("//*[@id='ExternalEmailAddress' and @data-val='true']");
        //private IWebElement txtInternalEmailAddressElem => ByXPath("//*[@id='InternalEmailAddress' and @data-val='true']");
        private IWebElement txtInternalEmailAddressElem => ByXPath(
            "/html/body/div[4]/div[1]/div[2]/div[2]/div[1]/div[2]/div[2]/div/div/form/div[6]/div[2]/div/input");
        private IWebElement txtMaximumRecipientsElem => ByXPath("//*[@id='MaximumRecipients']/input[@type='text']");

        private IWebElement ckbxMaximumRecipientsUnlimitedElem =>
            ByXPath("//*[@id='MaximumRecipients']//input[@type='checkbox']");
        private IWebElement ckbxMaximumReceiveSizeUnlimitedElem =>
            ByXPath("//*[@id='MaxReceiveSize']//input[@type='checkbox']");
        private IWebElement txtMaximumReceiveSizeElem => ByXPath("//*[@id='MaximumRecipients']/input[@type='text']");
        private IWebElement ckbxHideFromAddressList => ByXPath("//*[@id='HiddenFromAddressLists']");
        private IWebElement btnCreateMailContact => ByXPath("//button[@type = 'submit']");
        private IWebElement btnUserProfileElem => DriverContext.Driver.FindElement(By.CssSelector("[href*='#userprofile']"));
        private IWebElement txtGeneralProfileOfficeLocationElem => ById("GeneralProfile_OfficeLocation");
        private IWebElement txtGeneralProfileStreetAddressElem => ById("GeneralProfile_StreetAddress");
        private IWebElement txtGeneralProfileCityElem => ById("GeneralProfile_City");
        private IWebElement txtGeneralProfileZipCodeElem => ById("GeneralProfile_ZipCode");
        private IWebElement txtGeneralProfileJobTitleElem => ById("GeneralProfile_JobTitle");
        private IWebElement txtxGeneralProfileCompanyElem => ById("GeneralProfile_Company");
        private IWebElement btnAddManagerElem => ByXPath("//button[contains(@onclick, 'ExgMailboxManager.GetMailboxManager')]");
        private IWebElement txtGeneralProfileBusinessPhoneElem => ById("GeneralProfile_BusinessPhone");
        private IWebElement txtGeneralProfileFaxElem => ById("GeneralProfile_Fax");
        private IWebElement txtGeneralProfileHomePhoneElem => ById("GeneralProfile_HomePhone");
        private IWebElement txtGeneralProfilePagerElem => ById("GeneralProfile_Pager");
        private IWebElement txtGeneralProfileNotesElem => ById("GeneralProfile_Notes");
        private IWebElement txtGeneralProfileMobilePhoneElem => ById("GeneralProfile_MobilePhone");
        private IWebElement txtGeneralProfileDepartmentElem => ById("GeneralProfile_Department");
        private IWebElement dropdownCountryElem => ByXPath("//*[@id='select2-Country-container']");

        private IWebElement dropdownStateElem => ByXPath("//*[@id='GeneralProfile_ExistingState']");



        public string CreateMailContact(bool isSubOU, string contactName, string firstName, string lastName,
            string externalEmail, string internalEmail, bool isHiddenFromAddressList, string maximumRecipients,
            string maximumReceiveSize,string displayName, string country, string state, string officeLocation, string address, string city,
            string zipCode, string jobTitle, string company, string department, string managedBy, string businessPhone,
            string fax, string homePhone, string mobilePhone, string pager, string notes)
        {

            try
            {
                if (isSubOU == true)
                {
                    SelectSubOU(DriverContext.Driver, lnkSubOUElem, subOUWaitElem, subOUWaitBy);
                }

                txtContactNameElem.SendKeys(contactName);
                txtFirstNameElem.SendKeys(firstName);
                txtLastNameElem.SendKeys(lastName);
                txtExternalEmailAddressElem.SendKeys(externalEmail);

                if (!string.IsNullOrEmpty(internalEmail))
                {
                    MailAddress addr = new MailAddress(internalEmail);
                    string userName = addr.User;
                    string mailDomain = addr.Host;
                    txtInternalEmailAddressElem.SendKeys(userName);
                }

                if (isHiddenFromAddressList)
                {
                    if (!ckbxHideFromAddressList.Selected)
                    {
                        ckbxHideFromAddressList.Click();
                    }
                }

                if (maximumRecipients != null)
                {
                    txtMaximumRecipientsElem.Clear();
                    txtMaximumRecipientsElem.SendKeys(maximumRecipients);
                }

                if (maximumReceiveSize != null)
                {
                    txtMaximumReceiveSizeElem.Clear();
                    txtMaximumReceiveSizeElem.SendKeys(maximumReceiveSize);
                }

                if (displayName != null || country != null || state != null || officeLocation != null ||
                    address != null || city != null || zipCode != null || jobTitle != null || company != null ||
                    department != null || managedBy != null || businessPhone != null || fax != null ||
                    homePhone != null || mobilePhone != null || pager != null || notes != null)
                {
                    btnUserProfileElem.Click();

                    if (displayName != null)
                    {
                        txtDisplayNameElem.SendKeys(displayName);
                    }

                    //if (country != null)
                    //{
                    //    SetCountryAndState(country, state);
                    //}

                    if (officeLocation != null)
                    {
                        txtGeneralProfileOfficeLocationElem.Clear();
                        txtGeneralProfileOfficeLocationElem.SendKeys(officeLocation);
                    }

                    if (address != null)
                    {
                        txtGeneralProfileStreetAddressElem.Clear();
                        txtGeneralProfileStreetAddressElem.SendKeys(address);
                    }

                    if (city != null)
                    {
                        txtGeneralProfileCityElem.Clear();
                        txtGeneralProfileCityElem.SendKeys(city);
                    }

                    if (zipCode != null)
                    {
                        txtGeneralProfileZipCodeElem.Clear();
                        txtGeneralProfileZipCodeElem.SendKeys(zipCode);
                    }

                    if (jobTitle != null)
                    {
                        txtGeneralProfileJobTitleElem.Clear();
                        txtGeneralProfileJobTitleElem.SendKeys(jobTitle);
                    }

                    if (company != null)
                    {
                        txtxGeneralProfileCompanyElem.Clear();
                        txtxGeneralProfileCompanyElem.SendKeys(company);
                    }

                    if (department != null)
                    {
                        txtGeneralProfileDepartmentElem.Clear();
                        txtGeneralProfileDepartmentElem.SendKeys(department);
                    }

                    if (managedBy != null)
                    {
                        btnAddManagerElem.Click();
                        DriverContext.Driver.SwitchTo().Window(DriverContext.Driver.WindowHandles.Last());
                        DriverContext.Driver.FindElement(By.XPath("//input[contains(@id, '" + managedBy + "')]"))
                            .Click();
                        DriverContext.Driver
                            .FindElement(By.XPath("/html/body/div[2]/div/div[4]/div/div/form/div/div/div/button[2]"))
                            .Click();
                        DriverContext.Driver.SwitchTo().Window(DriverContext.Driver.WindowHandles.First());
                    }

                    if (businessPhone != null)
                    {
                        txtGeneralProfileBusinessPhoneElem.Clear();
                        txtGeneralProfileBusinessPhoneElem.SendKeys(businessPhone);
                    }

                    if (fax != null)
                    {
                        txtGeneralProfileFaxElem.Clear();
                        txtGeneralProfileFaxElem.SendKeys(fax);
                    }

                    if (homePhone != null)
                    {
                        txtGeneralProfileHomePhoneElem.Clear();
                        txtGeneralProfileHomePhoneElem.SendKeys(homePhone);
                    }

                    if (mobilePhone != null)
                    {
                        txtGeneralProfileMobilePhoneElem.Clear();
                        txtGeneralProfileMobilePhoneElem.SendKeys(mobilePhone);
                    }

                    if (pager != null)
                    {
                        txtGeneralProfilePagerElem.Clear();
                        txtGeneralProfilePagerElem.SendKeys(pager);
                    }

                    if (notes != null)
                    {
                        txtGeneralProfileNotesElem.Clear();
                        txtGeneralProfileNotesElem.SendKeys(notes);
                    }

                    Thread.Sleep(2000);
                    btnCreateMailContact.Click();
                    
                }
                
                return  GetPrompt(headerProgressElem, headerProgressElemBy, MessageContainer.DialogeContainer);
                
            }

            

            catch (Exception e)
            {
                return e.Message.Trim();
            }



            
        }

    }
}
