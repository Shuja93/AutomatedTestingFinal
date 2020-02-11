using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using HC10AutomationFramework.Helpers;

namespace TestRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            string userLevel;
            string browser;
            string module;
            if (args == null || args.Length < 3)
            {
                Console.WriteLine("Please add correct options");
            }
            else
            {
                userLevel = args[0];
                module = args[1];
                browser = args[2];
                
               

                string xmlFilePath = SetDirectory() + "\\AutomationFramework\\Config\\GlobalConfig.xml";
                XDocument xdoc = XDocument.Load(xmlFilePath);
                var node = xdoc.Descendants("RunSettings").FirstOrDefault();
                node.SetElementValue("Browser", browser);
                node.SetElementValue("Module", module);
                node.SetElementValue("UserLevel",userLevel);
                node.Save(xmlFilePath);
            }
            
        }

        private static string SetDirectory()
        {
            string dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string newDir = dir.Replace("\\TestRunner\\bin\\Debug\\netcoreapp3.1\\TestRunner.dll", "");
            return newDir;
        }
    }
}


