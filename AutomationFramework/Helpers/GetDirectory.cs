using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HC10AutomationFramework.Helpers
{
    public static class GetDirectory
    {
        public static string GetDir()
        {
            string dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
            dir = dir.Replace("\\HC10Test\\bin\\Debug\\HC10AutomationFramework.dll", "");
            //dir = dir.Replace("\\AutomatedTesting", "");
            
            

            return dir;
        }

        public static void SetDataDirectory()
        {
            string dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
            dir = dir.Replace("\\HC10Test\\bin\\Debug\\HC10AutomationFramework.dll", "");
            AppDomain.CurrentDomain.SetData("DataDirectory", dir);
            var test = AppDomain.CurrentDomain.GetData("DataDirectory");
        }
    }
}
