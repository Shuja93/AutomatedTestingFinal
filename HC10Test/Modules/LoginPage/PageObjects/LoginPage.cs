using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using System;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Config;
using HC10AutomationFramework.Helpers;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace HC10Test.PageObjects
{
    public class LoginPage : BasePage
    {

        private IWebElement txtUserName => DriverContext.Driver.FindElement(By.Id("LoginName"));
        private IWebElement txtPassword => DriverContext.Driver.FindElement(By.Id("Password"));
        private IWebElement btnLogin => DriverContext.Driver.FindElement(By.XPath("//button[2]"));
        private string path = "";
        _Application excel = new _Excel.Application();
        private Workbook wb;
        private Worksheet ws;

        public LoginPage()
        {
            string path = GetDirectory.GetDir() + "\\HC10Test\\Modules\\LoginPage\\Data\\Credentials.xlsx";
            string sheet = Settings.UserLevel;
            sheet.ToLower();
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }
        

        public void Login()
        {
            txtUserName.SendKeys(Convert.ToString(ws.Cells[2, 1].Value));
            txtPassword.SendKeys(Convert.ToString(ws.Cells[2, 2].Value));

            btnLogin.Click();
        }

        
        //public void SuccessStatus()
        // {
        // Console.WriteLine(errorMessage.Text);
        //}


    }
   
}
 