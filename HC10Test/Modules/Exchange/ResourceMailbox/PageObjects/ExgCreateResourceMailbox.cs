using HC10AutomationFramework.Helpers;
using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System.Net.Mail;
using System;
using System.Linq;
using HC10AutomationFramework.Enum;
using System.Threading;
using HC10AutomationFramework.Logs;



namespace HC10Test.PageObjects
{
     class ExgCreateResourceMailbox : BasePage
    {
        private IWebElement dropdownMailboxTypeElem => ById("MailboxType");


        //private IWebElement txtCountryElem => ByXPath("//*[contains(@class, 'select2-search__field')]");
        //private IWebElement dialogStatusMessageContainer => ByXPath("//*[@id='DialogStatusMessageContainer']/div");

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
        private IWebElement txtEmailPrefixElem => ByXPath("//input[@id = 'DisplayName' and @type = 'text' and @data-val-required = 'Please enter a valid username.']");

        private IWebElement btnCreateResourceMailboxElem =>
            ByXPath("//*[@id='SaveResoureceBtn']");
        private IWebElement  dropdownCountryElem => ByXPath("//*[@id='select2-Country-container']");

        private IWebElement dropdownStateElem => ByXPath("//*[@id='GeneralProfile_ExistingState']");


        public string CreateResourceMailbox(string resourceType, bool isSubOU, string email, bool isCR, string mailboxSize, string firstname, string lastName, string displayName, string country, string state, string officeLocation, string address, string city, string zipCode, string jobTitle, string company, string department, string businessPhone, string
            fax, string homePhone, string mobilePhone, string pager, string notes)
        {
            try
            {
                SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCreateResourceMailboxElem);

                MailAddress addr = new MailAddress(email);
                string userName = addr.User;
                string mailDomain = addr.Host;
                txtEmailPrefixElem.SendKeys(userName);


                if (resourceType == "Equipment")
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownMailboxTypeElem, "Equipment");
                }

                if (isSubOU == true)
                {
                    SelectSubOU(DriverContext.Driver, lnkSubOUElem, subOUWaitElem, subOUWaitBy);
                }


                if (isCR == true)
                {
                    SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, mailboxSize);
                }

                else
                {
                    SetResourceMailboxSize(mailboxSize);
                }


                if (firstname != null || lastName != null || displayName != null || country != null || state != null || officeLocation != null || address != null || city != null || zipCode != null || jobTitle != null || company != null || department != null || businessPhone != null || fax != null || homePhone != null || mobilePhone != null || pager != null || notes != null)
                {
                    btnUserProfileElem.Click();
                    if (firstname != null) { txtFirstNameElem.SendKeys(firstname); }
                    if (lastName != null) { txtLastNameElem.SendKeys(lastName); }
                    if (displayName != null) { txtDisplayNameElem.SendKeys(displayName); }

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

                }

                Thread.Sleep(2000);
                btnCreateResourceMailboxElem.Click();
                //PageRefresh(DriverContext.Driver);

                return GetPrompt(headerProgressElem,headerProgressElemBy,MessageContainer.DialogeContainer);

            }
            catch (Exception ex)
            {
                return ex.Message;

            }
            

            
        }

        void SetResourceMailboxSize(string mailboxSize)
        {
            if (ckbxMailboxSizeUnlimitedElem.Selected)
            {
                ckbxMailboxSizeUnlimitedElem.Click();
            }
            txtMailboxSizeElem.Clear();
            txtMailboxSizeElem.SendKeys(mailboxSize);
            if (ckbxIncomingSizeUnlimitedElem.Selected)
            {
                ckbxIncomingSizeUnlimitedElem.Click();
            }
            txtIncomingSize.Click();
            txtIncomingSize.Clear();
            txtIncomingSize.SendKeys(mailboxSize);
            if (ckbxOutgoingSizeUnlimited.Selected)
            {
                ckbxOutgoingSizeUnlimited.Click();
            }
            txtOutgoingSize.Clear();
            txtOutgoingSize.SendKeys(mailboxSize);
        }

        
            
        }







    }

