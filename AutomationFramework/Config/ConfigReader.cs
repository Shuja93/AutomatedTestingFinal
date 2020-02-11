using System;
using System.IO;
using System.Xml.XPath;
using HC10AutomationFramework.Base;
using HC10AutomationFramework.Helpers;

namespace HC10AutomationFramework.Config
{
    public class ConfigReader 
    {
        public static void SetFrameworkSettings()
        {
            XPathItem aut;
            XPathItem userLevel;
            XPathItem browsertype;
            XPathItem module;


            string dir = GetDirectory.GetDir();
            string strFileName = dir + "\\AutomationFramework\\Config\\GlobalConfig.xml";
            FileStream stream = new FileStream(strFileName, FileMode.Open);
            XPathDocument document = new XPathDocument(stream);
            XPathNavigator navigator = document.CreateNavigator();

            aut = navigator.SelectSingleNode("/RunSettings/AUT");
            browsertype = navigator.SelectSingleNode("/RunSettings/Browser");
            userLevel = navigator.SelectSingleNode("/RunSettings/UserLevel");
            module = navigator.SelectSingleNode("/RunSettings/Module");
           


            Settings.AUT = aut.Value;
            Settings.BrowserType = (BrowserType)System.Enum.Parse(typeof(BrowserType), browsertype.Value.ToString());
            Settings.UserLevel = userLevel.Value;
            Settings.Module = module.Value;
           

        }

    }
}
