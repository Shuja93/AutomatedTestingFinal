using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HC10AutomationFramework.TestTracker
{
    public static class TestTracker
    {
        public static IDictionary<string, string> mailboxStatus = new Dictionary<string, string>();
        public static IDictionary<string, string> distributionListStatus = new Dictionary<string, string>();
        public static IDictionary<string, string> mailContactStatus = new Dictionary<string, string>();
        public static IDictionary<string, string> resourceMailboxStatus = new Dictionary<string, string>();
        public static IDictionary<string, string> publicFolderStatus = new Dictionary<string, string>();
    }
}
