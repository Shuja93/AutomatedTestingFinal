using HC10AutomationFramework.Helpers;
using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System.Net.Mail;
using System;
using System.Linq;
using HC10AutomationFramework.Enum;
using System.Threading;
using HC10AutomationFramework.Logs;
using Microsoft.Office.Interop.Excel;
using HC10AutomationFramework.Resources;


namespace HC10Test.PageObjects
{
    class ExgCreateMailbox : BasePage
    {
        private IWebElement dropdownMailboxTypeElem => ById("MailboxType");
        private IWebElement radioNewUserElem => ByXPath("//*[@id='CreateNewUser']");
        private IWebElement radioExistingUserElem => ByXPath("//*[@id='UseExistingUser']");
        private IWebElement txtEmailPrefixElem => ByXPath("//*[@id='notlinkeduser']/div/div[1]/div/div[1]/div/input");
        private IWebElement dropdownMailDomainElem => ByXPath("MailboxType");
        private IWebElement txtPasswordElem => ByXPath("//*[@id='Password']");
        private IWebElement txtConfirmPasswordElem => ByXPath("//*[@id='ConfirmPassword']");

        //private IWebElement txtCountryElem => ByXPath("//*[contains(@class, 'select2-search__field')]");
        //private IWebElement dialogStatusMessageContainer => ByXPath("//*[@id='DialogStatusMessageContainer']/div");
        private IWebElement ckbxMailboxUnlimitedElem => ByXPath("//*[@id='MailboxSize']/label/input");
        private IWebElement ckbxChangePasswordElem => ByXPath("//*[@id='ResetPasswordOnNextLogon']");
        //private IWebElement btnCreateMailboxElem => ByXPath("//*[@id='SaveBtn']");
        private IWebElement btnChooseExistingElem => ByXPath("//*[@id='existinguser']//button");
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
        private IWebElement  dropdownCountryElem => ByXPath("//*[@id='select2-Country-container']");

        private IWebElement dropdownStateElem => ByXPath("//*[@id='GeneralProfile_ExistingState']");
        //private IWebElement btnCloseDialogueBox =>
            //ByXPath("//button[contains(text(), 'Cancel') and contains(@onclick, 'Layout.CloseModelDialog()')]");

        public string CreateMailbox(string mailboxType, bool isSubOU, bool newUser, string email, string mailboxPassword, bool isCR, string mailboxSize, bool passwordChange, string firstname, string lastName, string displayName, string country, string state, string officeLocation, string address, string city, string zipCode, string jobTitle, string company, string department, string managedBy, string businessPhone, string
            fax, string homePhone, string mobilePhone, string pager, string notes)
        {
            try
            {
                if (!SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCreateMailboxElem))
                {
                    return ErrorDescriptions.CreateButtonTimeout;
                }

                if (mailboxType == "Shared")
                {
                    if (!SeleniumHelperMethods.SelectDropDownValue(dropdownMailboxTypeElem, "Shared Mailbox"))
                    {
                        return ErrorDescriptions.ErrorDropdownMenu;
                    }

                }

                if (isSubOU == true)
                {
                    if (!SelectSubOU(DriverContext.Driver, lnkSubOUElem, subOUWaitElem,
                        subOUWaitBy))
                    {
                        return ErrorDescriptions.ErrorSubOuNotFound;
                    }
                }
                

                if (newUser == true)
                {
                    MailAddress addr = new MailAddress(email);
                    string userName = addr.User;
                    string mailDomain = addr.Host;
                    txtEmailPrefixElem.SendKeys(userName);
                    //SeleniumHelperMethods.SelectDropDownValue(dropdownMailDomain, mailDomain);
                    txtPasswordElem.SendKeys(mailboxPassword);
                    txtConfirmPasswordElem.SendKeys(mailboxPassword);
                }

                else
                {
                    radioExistingUserElem.Click();
                    btnChooseExistingElem.Click();
                    SelectExistingObject(DriverContext.Driver, email);

                }

                if (isCR == true)
                {
                    if (!SeleniumHelperMethods.SelectDropDownValue(dropdownCRElem, mailboxSize))
                    {
                        return ErrorDescriptions.ErrorSelectingCR;
                    }
                }

                else
                {
                    SetMailboxSize(mailboxSize);
                }

                if (passwordChange == true)
                {
                    ckbxChangePasswordElem.Click();
                }

                if (firstname != null || lastName != null || displayName != null || country != null || state != null || officeLocation != null || address != null || city != null || zipCode != null || jobTitle != null || company != null || department != null || managedBy != null || businessPhone != null || fax != null || homePhone != null || mobilePhone != null || pager != null || notes != null)
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

                    if (managedBy != null)
                    {
                        try
                        {
                            btnAddManagerElem.Click();
                            DriverContext.Driver.SwitchTo().Window(DriverContext.Driver.WindowHandles.Last());
                            DriverContext.Driver.FindElement(By.XPath("//input[contains(@id, '" + managedBy + "')]")).Click();
                            DriverContext.Driver.FindElement(By.XPath("/html/body/div[2]/div/div[4]/div/div/form/div/div/div/button[2]")).Click();
                            DriverContext.Driver.SwitchTo().Window(DriverContext.Driver.WindowHandles.First());

                        }
                        catch (Exception )
                        {
                            return ErrorDescriptions.ErrorManagerNotAdded;
                        }
                   
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
                btnCreateMailboxElem.Click();
                //PageRefresh(DriverContext.Driver);

                return GetPrompt(headerProgressElem,headerProgressElemBy,MessageContainer.DialogeContainer);

            }
            catch (Exception ex)
            {
                return ex.Message;

            }
            

            
        }

        //public void CloseDialogueBox()
        //{
        //    SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCloseDialogueBox);
        //    Thread.Sleep(2000);
        //    btnCloseDialogueBox.Click();
        //}

        
            
        }







    }

