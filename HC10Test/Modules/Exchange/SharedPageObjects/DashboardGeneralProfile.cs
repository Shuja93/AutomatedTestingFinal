using OpenQA.Selenium;
using HC10AutomationFramework.Base;
using System;
using OpenQA.Selenium.Support.UI;
using System.Linq;
using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Resources;

namespace HC10Test.PageObjects
{
    public class DashboardGeneralProfile : BasePage
    {

        private IWebElement btnMailBoxGenPropSaveElem => ByXPath("//button[@type = 'submit']");
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
        private IWebElement listManager => ByXPath("//*[@id='managerList']//p");


        public string VerifyGeneralProperties(string firstname, string lastName, string displayName, string country, string state, string
                    officeLocation, string address, string city, string zipCode, string jobTitle, string company, string department, string managedBy, string businessPhone, string
                    fax, string homePhone, string mobilePhone, string pager, string notes)
        {
            try
            {
                expectedProperties.Add("FirstName", firstname);
                expectedProperties.Add("LastName", lastName);
                expectedProperties.Add("DisplayName", displayName);
                //expectedProperties.Add("Country", country);
                //expectedProperties.Add("State", state);
                expectedProperties.Add("Office Location", officeLocation);
                expectedProperties.Add("Address", address);
                expectedProperties.Add("City", city);
                expectedProperties.Add("Zip Code", zipCode);
                expectedProperties.Add("Job Title", jobTitle);
                expectedProperties.Add("Company", company);
                expectedProperties.Add("Department", department);
                if (!string.IsNullOrEmpty(managedBy))
                {
                    expectedProperties.Add("ManagedBy", managedBy);
                }
                
                expectedProperties.Add("Business Phone", businessPhone);
                expectedProperties.Add("Fax", fax);
                expectedProperties.Add("Home Phone", homePhone);
                expectedProperties.Add("Mobile Phone", mobilePhone);
                expectedProperties.Add("Pager", pager);
                expectedProperties.Add("Notes", notes);


                actualProperties.Add("FirstName", Convert.ToString(txtFirstNameElem.GetAttribute("value")));
                actualProperties.Add("LastName", Convert.ToString(txtLastNameElem.GetAttribute("value")));
                actualProperties.Add("DisplayName", Convert.ToString(txtDisplayNameElem.GetAttribute("value")));
                //actualProperties.Add("Country", dropdownCountryElem.Text);
                var selectElement = new SelectElement(dropdownStateElem);
                string _state = selectElement.SelectedOption.Text;
                //actualProperties.Add("State", _state);
                actualProperties.Add("Office Location", Convert.ToString(txtGeneralProfileOfficeLocationElem.GetAttribute("value")));
                actualProperties.Add("Address", Convert.ToString(txtGeneralProfileStreetAddressElem.GetAttribute("value")));
                actualProperties.Add("City", Convert.ToString(txtGeneralProfileCityElem.GetAttribute("value")));
                actualProperties.Add("Zip Code", Convert.ToString(txtGeneralProfileZipCodeElem.GetAttribute("value")));
                actualProperties.Add("Job Title", Convert.ToString(txtGeneralProfileJobTitleElem.GetAttribute("value")));
                actualProperties.Add("Company", Convert.ToString(txtxGeneralProfileCompanyElem.GetAttribute("value")));
                actualProperties.Add("Department", Convert.ToString(txtGeneralProfileDepartmentElem.GetAttribute("value")));
                if (!string.IsNullOrEmpty(managedBy))
                {
                    actualProperties.Add("ManagedBy", Convert.ToString(listManager.Text));
                }
                actualProperties.Add("Business Phone", Convert.ToString(txtGeneralProfileBusinessPhoneElem.GetAttribute("value")));
                actualProperties.Add("Fax", Convert.ToString(txtGeneralProfileFaxElem.GetAttribute("value")));
                actualProperties.Add("Home Phone", Convert.ToString(txtGeneralProfileHomePhoneElem.GetAttribute("value")));
                actualProperties.Add("Mobile Phone", Convert.ToString(txtGeneralProfileMobilePhoneElem.GetAttribute("value")));
                actualProperties.Add("Pager", Convert.ToString(txtGeneralProfilePagerElem.GetAttribute("value")));
                actualProperties.Add("Notes", Convert.ToString(txtGeneralProfileNotesElem.GetAttribute("value")));

                return CompareLists(expectedProperties, actualProperties);
                
            }

            catch (Exception ex)
            {
                return ex.Message;
            }

        }

        public string VerifyGeneralProperties(string firstname, string lastName, string displayName, string country, string state, string
            officeLocation, string address, string city, string zipCode, string jobTitle, string company, string department, string businessPhone, string
            fax, string homePhone, string mobilePhone, string pager, string notes)
        {
            try
            {
                expectedProperties.Add("FirstName", firstname);
                expectedProperties.Add("LastName", lastName);
                expectedProperties.Add("DisplayName", displayName);
                //expectedProperties.Add("Country", country);
                //expectedProperties.Add("State", state);
                expectedProperties.Add("Office Location", officeLocation);
                expectedProperties.Add("Address", address);
                expectedProperties.Add("City", city);
                expectedProperties.Add("Zip Code", zipCode);
                expectedProperties.Add("Job Title", jobTitle);
                expectedProperties.Add("Company", company);
                expectedProperties.Add("Department", department);
                expectedProperties.Add("Business Phone", businessPhone);
                expectedProperties.Add("Fax", fax);
                expectedProperties.Add("Home Phone", homePhone);
                expectedProperties.Add("Mobile Phone", mobilePhone);
                expectedProperties.Add("Pager", pager);
                expectedProperties.Add("Notes", notes);


                actualProperties.Add("FirstName", Convert.ToString(txtFirstNameElem.GetAttribute("value")));
                actualProperties.Add("LastName", Convert.ToString(txtLastNameElem.GetAttribute("value")));
                actualProperties.Add("DisplayName", Convert.ToString(txtDisplayNameElem.GetAttribute("value")));
                //actualProperties.Add("Country", dropdownCountryElem.Text);
                var selectElement = new SelectElement(dropdownStateElem);
                string _state = selectElement.SelectedOption.Text;
                //actualProperties.Add("State", _state);
                actualProperties.Add("Office Location", Convert.ToString(txtGeneralProfileOfficeLocationElem.GetAttribute("value")));
                actualProperties.Add("Address", Convert.ToString(txtGeneralProfileStreetAddressElem.GetAttribute("value")));
                actualProperties.Add("City", Convert.ToString(txtGeneralProfileCityElem.GetAttribute("value")));
                actualProperties.Add("Zip Code", Convert.ToString(txtGeneralProfileZipCodeElem.GetAttribute("value")));
                actualProperties.Add("Job Title", Convert.ToString(txtGeneralProfileJobTitleElem.GetAttribute("value")));
                actualProperties.Add("Company", Convert.ToString(txtxGeneralProfileCompanyElem.GetAttribute("value")));
                actualProperties.Add("Department", Convert.ToString(txtGeneralProfileDepartmentElem.GetAttribute("value")));
                actualProperties.Add("Business Phone", Convert.ToString(txtGeneralProfileBusinessPhoneElem.GetAttribute("value")));
                actualProperties.Add("Fax", Convert.ToString(txtGeneralProfileFaxElem.GetAttribute("value")));
                actualProperties.Add("Home Phone", Convert.ToString(txtGeneralProfileHomePhoneElem.GetAttribute("value")));
                actualProperties.Add("Mobile Phone", Convert.ToString(txtGeneralProfileMobilePhoneElem.GetAttribute("value")));
                actualProperties.Add("Pager", Convert.ToString(txtGeneralProfilePagerElem.GetAttribute("value")));
                actualProperties.Add("Notes", Convert.ToString(txtGeneralProfileNotesElem.GetAttribute("value")));

                return CompareLists(expectedProperties, actualProperties);

            }

            catch (Exception ex)
            {
                return ex.Message;
            }

        }

        public string SetGeneralProperties(string firstname, string lastName, string displayName, string country, string state, string officeLocation, string address, string city, string zipCode, string jobTitle, string company, string department, string managedBy, string businessPhone, string fax, string homePhone, string mobilePhone, string pager, string notes)
        {
            try
            {
                if (firstname != null)
                {
                    txtFirstNameElem.Clear();
                    txtFirstNameElem.SendKeys(firstname);
                }

                if (lastName != null)
                {
                    txtLastNameElem.Clear();
                    txtLastNameElem.SendKeys(lastName);
                }

                if (displayName != null)
                {
                    txtDisplayNameElem.Clear();
                    txtDisplayNameElem.SendKeys(displayName);
                }

                //if (country != null)
                //{
                //    SetCountryAndState(country,state);
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

                if (!string.IsNullOrEmpty(managedBy))
                {
                    try
                    {
                        SetManager(DriverContext.Driver, managedBy);
                    }
                    catch (NoSuchElementException)
                    {
                        return ErrorDescriptions.ErrorUserNotFound;
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
                btnMailBoxGenPropSaveElem.Click();
                return GetPrompt(headerProgressElem, headerProgressElemBy, MessageContainer.ToastContainer);

            }
            catch (Exception ex)
            {
                return ex.Message;

            }

        }

        private void SetManager(IWebDriver driver, string managedBy) 
        {
           
            btnAddManagerElem.Click();
            driver.SwitchTo().Window(DriverContext.Driver.WindowHandles.Last());
            driver.FindElement(By.XPath("//input[contains(@id, '" + managedBy + "')]")).Click();
            driver.FindElement(By.XPath("/html/body/div[2]/div/div[4]/div/div/form/div/div/div/button[2]")).Click();
            driver.SwitchTo().Window(DriverContext.Driver.WindowHandles.First());
        }
    }
}
