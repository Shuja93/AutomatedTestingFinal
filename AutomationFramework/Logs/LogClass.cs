using System;
using System.IO;

namespace HC10AutomationFramework.Logs
{
    public class LogClass
    {
        
        
        public static void AppendLogs(Exception ex)
        {
            using (StreamWriter w = File.AppendText("C:\\Users\\ShujaPC\\source\\repos\\AutomationFramework\\AutomatedTesting\\AutomationFramework\\Logs\\Logfile.txt"))
            {
                
                w.WriteLine(ex);
            }
        }

        public static void AppendLogs(string message)
        {
            using (StreamWriter w = File.AppendText("C:\\Users\\ShujaPC\\source\\repos\\AutomationFramework\\AutomatedTesting\\AutomationFramework\\Logs\\Logfile.txt"))
            {
                w.WriteLine(message);
            }
        }
    }
}
