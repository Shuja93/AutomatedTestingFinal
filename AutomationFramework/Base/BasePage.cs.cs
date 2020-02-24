using HC10AutomationFramework.Enum;
using HC10AutomationFramework.Helpers;
using OpenQA.Selenium;
using HC10AutomationFramework.Logs;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Linq;
using HC10AutomationFramework.Extensions;
using HC10AutomationFramework.Resources;
using OpenQA.Selenium.Support.UI;

namespace HC10AutomationFramework.Base
{
    public abstract class BasePage : TestBase
    {
        //public bool VerifyUserLevel(T)
        public  bool SelectSubOU(IWebDriver driver, IWebElement subOU, IWebElement loadwaitElement, string loadwaitString)
        {
            try
            {
                SetDriverTime(30);
                subOU.Click();
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                driver.FindElement(By.ClassName("jstree-anchor")).Click();
                driver.FindElement(By.XPath("//*[@id='selectFileBtn']")).Click();
                driver.SwitchTo().Window(driver.WindowHandles.First());
                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }

        public  void SelectExistingObject(IWebDriver driver, string id)
        {
            
            var windows = driver.WindowHandles;
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            driver.FindElement(By.XPath("//input[@type = 'radio' and contains(@id, '" + id + "')] ")).Click();
            driver.FindElement(By.XPath("/html/body/div[2]/div/div[2]/div/div/form/div/div/div/button[2]")).Click();
            driver.SwitchTo().Window(windows[0]);

        }

        //public List<string> issueKey = new List<string>();
        public IDictionary<string, string> expectedProperties = new Dictionary<string, string>();
        public IDictionary<string, string> actualProperties = new Dictionary<string, string>();


        public string GetPrompt(IWebElement loadWaitElement, string loadWaitString, string containerType)
        {
            SetDriverTime(2);
            try
            {
                string prompt;
                SeleniumHelperMethods.LoadWait(DriverContext.Driver, loadWaitElement, loadWaitString);
                try
                {
                        prompt = containerType == MessageContainer.DialogeContainer
                        ? dialogueContainerElem.Text
                        : toastContainer.Text;

                }
                catch (NoSuchElementException)
                {
                    prompt = toastContainer.Text;
                    return prompt;
                }

                prompt = prompt.Trim();
                if (containerType == MessageContainer.DialogeContainer) { prompt = prompt.Substring(10); }
                SetDriverTime(30);
                return prompt;
            }
            

            catch (Exception e)
            {
                return e.Message.Trim();
            }

        }

        public string VerifyResult(string expectedMessage, string actualMessage)
        {
            if (!actualMessage.Contains(expectedMessage))
            {
                return TestStatus.Failed;
            }

            return TestStatus.Success;
        }

        //revisit
        public bool InternalErrorPresence(IWebDriver driver)
        {
            SetDriverTime(2);
            try
            {
                var errorPresence = driver.FindElement(By.CssSelector("error-desc"));
                SetDriverTime(30);
                return true;
            }

            catch (NoSuchElementException)
            {
                SetDriverTime(30);
                return false;
            }
            
        }

        public void SetMailboxSize(string size)
        {
            txtMailboxSizeElem.Clear();
            txtMailboxSizeElem.SendKeys(size);
            if (Convert.ToString(DriverContext.Driver
                    .FindElement(By.XPath("//*[@id='MaxIncomingMsgSize']/input[@type='hidden']"))
                    .GetAttribute("value")) == "-1")
            {
                ckbxIncomingSizeUnlimitedElem.Click();
            }

            txtIncomingSize.Clear();
            txtIncomingSize.SendKeys(size);
            if (Convert.ToString(DriverContext.Driver
                    .FindElement(By.XPath("//*[@id='MaxOutgoingMsgSize']/input[@type='hidden']"))
                    .GetAttribute("value")) == "-1")
            {
                ckbxOutgoingSizeUnlimited.Click();
            }

            txtOutgoingSize.Clear();
            txtOutgoingSize.SendKeys(size);
            txtProhibitSendAtElem.Clear();
            txtProhibitSendAtElem.SendKeys(size);
            txtWarnAtElem.Clear();
            txtWarnAtElem.SendKeys(size);
        }

        public static void PageRefresh(IWebDriver driver)
        {
            Thread.Sleep(5000);
            driver.Navigate().Refresh();
            Thread.Sleep(2000);
        }

        public string AddUsersinNewWindows(IWebDriver driver, string usersList)
        {
            try
            {
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                var users = usersList.Split(' ');
                foreach (string user in users)
                {
                    driver.FindElement(By.XPath(
                            "//input[@type = 'checkbox' and contains(@value, '" + user + "')]"))
                        .Click();
                }

                driver.FindElement(By.XPath("//button[contains(text(),'Add')]")).Click();
                driver.FindElement(By.XPath("//button[contains(text(),'Select Objects')]")).Click();
                driver.SwitchTo().Window(driver.WindowHandles.First());
                return "success";
            }
            catch (NoSuchElementException)
            {
                return ErrorDescriptions.ErrorAddingUserinPopUp;
            }
            catch (Exception e)
            {
                return e.Message.Trim();
            }

        }

        public string AddMembersInDl(IWebDriver driver, string userList)
        {
            try
            {
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                var members = userList.Split(' ');
                foreach (var member in members)
                {
                    var _member = member.Split(';');
                    SeleniumHelperMethods.SelectDropDownValue(driver.FindElement(By.XPath("//*[@id='RecipientType']")),_member[0]);
                    driver.FindElement(By.XPath(
                            "//input[@type = 'checkbox' and contains(@value, '" + _member[1] + "')]"))
                        .Click();
                    driver.FindElement(By.XPath("//button[contains(text(),'Add')]")).Click();
                }

                
                driver.FindElement(By.XPath("//button[contains(text(),'Select Objects')]")).Click();
                driver.SwitchTo().Window(driver.WindowHandles.First());
                return "success";
            }
            catch (NoSuchElementException)
            {
                return ErrorDescriptions.ErrorAddingUserinPopUp;
            }
            catch (Exception e)
            {
                return e.Message.Trim();
            }
        }

        public string VerifyUsersInPermissions(IWebDriver driver, string userList, string divId)
        {
            Thread.Sleep(2000);
            userList = userList.Trim();
            string status = TestStatus.Success;
            try
            {
                var listLength =
                    driver.FindElements(By.XPath("//*[@id='" + divId + "']//*[@id='divSelectedItems']/ul/li"));
                var users = userList.Split(' ');
                if (listLength.Count != users.Length)
                {
                    return ErrorDescriptions.ErrorUserCount;
                }

                foreach (string user in users)
                {
                    string userPresence = driver
                        .FindElement(By.XPath("//*[@id='" + divId +
                                              "']//*[@id='divSelectedItems']/ul/li[contains(@id , '" + user +
                                              "')]/input")).GetAttribute("value");
                    if (user != userPresence)
                    {
                        status = TestStatus.Failed;
                    }
                }

                return status;
            }

            catch (NoSuchElementException)
            {
                return ErrorDescriptions.ErrorUserNotFound;
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        public void AddForwardingPopUp(IWebDriver driver, string user, string ou, string exchangeObject)
        {
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            var selectOU = new SelectElement(dropdownOU);
            selectOU.SelectByText(ou);
            var selectObject = new SelectElement(dropdownExchangeObjects);
            selectObject.SelectByText(exchangeObject);
            driver.FindElement(By.Id(user)).Click();
            driver.FindElement(By.XPath("//button[contains(@onclick, 'SelectADUser')]"))
                .Click();
            driver.SwitchTo().Window(driver.WindowHandles.First());
        }

        public void ClickPermissionsSaveButton(IWebDriver driver, string divID)
        {
            driver.FindElement(By.XPath("//*[@id='" + divID + "']//button[contains(@type,'submit')]")).Click();
        }

        public void ClickSendersForTheList(IWebDriver driver, string divID)
        {
            driver.FindElement(By.XPath("//*[@id='" + divID + "']//*[@id='SelectedSender']")).Click();
        }

        public string CompareLists(IDictionary<string, string> expectedProperties,
            IDictionary<string, string> actualProperties)
        {
            List<string> issueKey = new List<string>();
            foreach (var key in expectedProperties.Keys)
            {
                if (!expectedProperties[key].Equals(actualProperties[key]))
                {
                    issueKey.Add(key);
                }
            }

            return string.Join(";", issueKey);
        }

        public static void SetDriverTime(int time)
        {
            DriverContext.Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(time);
        }

        public void SetCountryAndState(string country, string state)
        {

            //dropdownCountryElem.Click();
            //txtCountryElem.SendKeys(country);
            //txtCountryElem.SendKeys(Keys.Enter);

            //if (DriverContext.Driver.FindElement(By.XPath("//*[@id='otherSpan']")).GetAttribute("style")
            //    .Contains("display: none"))
            //{
            //    try
            //    {
            //        if (DriverContext.Driver.FindElement(By.XPath("//*[@id='existingSpan'']")).GetAttribute("display")=="none") { txtNewStateElem.SendKeys(state); }
            //        else 
            //        {
            //            SetDriverTime(2);
            //            var selectState = new SelectElement(dropdownStateElem);
            //            selectState.SelectByValue(state);
            //        }
                    
            //    }
            //    catch (Exception)
            //    {
            //        btnAddStateElem.Click();
                    
            //    }
            //}
            //txtNewStateElem.SendKeys(state);
            //SetDriverTime(30);
        }
        public static void LogOut(IWebElement btnArrowToggleElem, IWebElement btnLogoutElem)
        {

            btnArrowToggleElem.Click();
            btnLogoutElem.Click();
        }
        public void CloseDialogueBox()
        {
            SeleniumHelperMethods.WaitExpectedConditionsClickable(DriverContext.Driver, btnCloseDialogueBox);
            Thread.Sleep(2000);
            btnCloseDialogueBox.Click();
        }

        public void OpenAdvancedProperties()
        {
            lnkadvancedPropertiesElem.ClickWithWait("spinner");
        }

        public void OpenEmailAddress()
        {
            lnkEmailAddressElem.ClickWithWait("spinner");
        }

        public void OpenSendOnBehalf()
        {
            lnkSendOnBehalfElem.ClickWithWait("spinner");
        }

        public void OpenFullAccessPermissions()
        {
            lnkFullAccessPermissionsElem.ClickWithWait("spinner");
        }

        public void OpenSendAsPermissions()
        {
            lnkSenAsPermissionsElem.ClickWithWait("spinner");
        }

        public void OpenAcceptedSenders()
        {
            lnkAcceptedSendersElem.ClickWithWait("spinner");
        }

        public void OpenRejectedSenders()
        {
            lnkRejectedSendersElem.ClickWithWait("spinner");
        }

        public void OpenForwarding()
        {
            lnkForwardingElem.ClickWithWait("spinner");
        }

        public void SetCheckBox(IWebElement ckbxElement, bool ckbxStatus)
        {
            if (ckbxStatus == true)
            {
                if (!ckbxElement.Selected)
                {
                    ckbxElement.Click();
                }
            }
            else
            {
                if (ckbxElement.Selected)
                {
                    ckbxElement.Click();
                }
            }
        }

    }

}

