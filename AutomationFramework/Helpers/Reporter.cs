using System;
using System.IO;
using System.Linq;
using System.Text;
using HC10AutomationFramework.Config;
using HC10AutomationFramework.Logs;


namespace HC10AutomationFramework.Helpers
{
    public class ReporterClass
    {
        private static string fileName;
        private static string file;
        private string userLevel = Settings.UserLevel;
        private string browser = Convert.ToString(Settings.BrowserType);
        private string module = Settings.Module;
        string columnName = "Module,User Role,Test Name,Test Description,OU Name,Object Name,Email,Test Parameters,Test Status,Message Details";
        
        
        
        
        
        public void CreateReporterFile()
        {
            string dir = GetDirectory.GetDir();
            
            fileName = userLevel.ToLower() + "_" + module.ToLower()+ "_" + browser.ToLower()+ ".csv";
            file = dir + "\\HC10Test\\Reports\\" + fileName;
            StringBuilder csvContent = new StringBuilder();
            csvContent.AppendLine(columnName);
            File.AppendAllText(file,csvContent.ToString());

        }

        

        public static void Reporter(string module,string userRole,string testName, string testDesc,string ou, string exchangeObject,string email, string testParam,string testStatus, string errorDetails) 
        {
            try
            {
                using (StreamWriter w = File.AppendText(file))
                {
                    w.WriteLine(module + "," + userRole + "," + testName + "," + testDesc + "," + ou + "," + exchangeObject + "," + email + "," + testParam + "," + testStatus + "," + errorDetails);
                }
            }
            catch (Exception e)
            {
                LogClass.AppendLogs(e);
            }
            
        }
    }
}
