﻿using System;
using System.IO;

namespace HC10AutomationFramework.Logs
{
    public class LogClass
    {
        private static string _logFileName = string.Format("{0:yyyymmddhhmmss}", DateTime.Now);
        private static StreamWriter _streamw = null;

        //Create a file which can store the log information
        public static void CreateLogFile()
        {
            string dir = @"C:\Users\Shuja\source\repos\AutomationFramework\AutomatedTesting\HC10Test\Log\Logfile.txt";
            if (Directory.Exists(dir))
            {
                _streamw = File.AppendText(dir + _logFileName + ".log");
            }
            else
            {
                Directory.CreateDirectory(dir);
                _streamw = File.AppendText(dir + _logFileName + ".log");
            }
        }



        //Create a method which can write the text in the log file
        //public static void Write(string logMessage)
        //{
        //    _streamw.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
        //    _streamw.WriteLine("    {0}", logMessage);
        //    _streamw.Flush();
        //}


        public static void AppendLogs(Exception ex)
        {
            using (StreamWriter w = File.AppendText(@"C:\Users\Shuja\source\repos\AutomationFramework\AutomatedTesting\HC10Test\Log\Logfile.txt"))
            {
                
                w.WriteLine(ex);
            }
        }

        public static void AppendLogs(string message)
        {
            using (StreamWriter w = File.AppendText(@"C:\Users\Shuja\source\repos\AutomationFramework\AutomatedTesting\HC10Test\Log\Logfile.txt"))
            {
                w.WriteLine(message);
            }
        }
    }
}
