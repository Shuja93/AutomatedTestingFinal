using System;
using System.Diagnostics;
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



                try
                {
                    SetXml(userLevel, module, browser);
                    ExecuteTestFile(module);
                    ConvertCSVFile(userLevel, module, browser);
                }
                catch (Exception e) 
                {
                    Console.WriteLine(e.Message);
                }
                
            }
            
        }

        private static string SetDirectory()
        {
            string dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string newDir = dir.Replace("\\TestRunner\\bin\\Debug\\netcoreapp3.1\\TestRunner.dll", "");
            return newDir;
        }

        private static void ExecuteTestFile(string testCategory) 
        {
            string filePath = SetDirectory()  + @"\HC10Test\bin\Debug\HC10Test.dll";
            string arguments = filePath + String.Format(" /TestCaseFilter:\"TestCategory={0}\"",testCategory);
            string vsConsolePath = "\"C:\\Program Files (x86)\\Microsoft Visual Studio\\2019\\Enterprise\\Common7\\IDE\\Extensions\\TestPlatform\\vstest.console.exe\"";
           
            var proc = new Process();
            proc.StartInfo.FileName = vsConsolePath;
            proc.StartInfo.Arguments = arguments;
            proc.Start();
            proc.WaitForExit();
            var exitCode = proc.ExitCode;
            proc.Close();

        }
        private static void ConvertCSVFile(string userLevel, string module, string browser)
        {
            string csvFileName = userLevel.ToLower() + "_" + module.ToLower() + "_" + browser.ToLower() + ".csv";
            string csvFilePath = SetDirectory() + @"\HC10Test\Reports\" + csvFileName;
            string reportFileName = userLevel.ToLower() + "_" + module.ToLower() + "_" + browser.ToLower() + ".html";
            string reportFilePath = SetDirectory() + @"\HC10Test\Reports\" + reportFileName;
            string command = "/C csvtotable -c \"HC10 Testing Report\" " + csvFilePath +" "+ reportFilePath +" -e";

            Process.Start("CMD.exe", command);
        

        }

        private static void SetXml(string userLevel, string module, string browser) 
        {
            string xmlFilePath = SetDirectory() + "\\AutomationFramework\\Config\\GlobalConfig.xml";
            XDocument xdoc = XDocument.Load(xmlFilePath);
            var node = xdoc.Descendants("RunSettings").FirstOrDefault();
            node.SetElementValue("Browser", browser);
            node.SetElementValue("Module", module);
            node.SetElementValue("UserLevel", userLevel);
            node.Save(xmlFilePath);
        }
    }
}


