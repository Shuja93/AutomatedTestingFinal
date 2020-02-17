using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HC10AutomationFramework.Enum
{
    public class ExchangeMessages
    {
        public static string CreateMailbox => "Exchange mailbox created successfully.";
        public static string DisableMailbox => "Exchange mailbox disabled successfully.";
        public static string EnableMailbox => "Exchange mailbox enabled successfully.";
        public static string UpdateMailboxGeneralProperties => "User general profile settings updated successfully.";
        public static string UpdateRetentionPolicy => "User mailbox retention policy updated successfully.";
        public static string AddEmailAddress => "Mailbox email address added successfully.";
        public static string AddSendOnBehalfUsers => "Send on behalf users added successfully.";
        public static string AddFullAccessPermissions => "Mailbox permission settings updated successfully.";
        public static string AddSendAsPermissions => "Mailbox permission settings updated successfully.";
        public static string AddAcceptedUsers => "Accepted senders added successfully.";
        public static string AddRejectedSenders => "Rejected senders added successfully.";
        public static string AddForwarding => "Mailbox forwarding address updated successfully.";
        public static string AddAutomaticReply => "Mailbox out of office configurations set successfully.";

        public static string AddArchive => "Archive mailbox added successfully.";
        public static string CreateDL => "Exchange distribution list added successfully.";
        public static string AddDlMembers => "Exchange distribution list members updated successfully.";
        public static string AddDlAdmins => "Exchange distribution list administrators updated successfully.";

        public static string UpdateDlAdvanceProperties => "Distribution list properties updated successfully.";
        public static string AddDlEmailAddress => "Email address added successfully.";
        public static string CreateMailContact => "Success: Exchange mail contact added successfully.";

        public static string AddMailContactEmailAddress => "Email address added successfully.";

        public static string CreateResourceMailbox => "Success: Resource mailbox created successfully.";
        public static string AddResourceEmailAddress => "Email address added successfully.";

        public static string CreatePublicFolder => "Success: Public folder added successfully.";

        public static string UpdateUserMailboxAdvanceProperties =>
            "User mailbox advanced properties updated successfully.";

        public static string AddSendAsPermissionDl => "Distribution list permission settings updated successfully.";
        public static string AddPublicFolderEmailAddress => "Public folder email address added successfully.";

        public static string AddPublicFolderForwardingAddress =>
            "Public folder forwarding address(es) updated successfully.";

        public static string UpdateMailContactAdvanceProperties => "Mail contact properties updated successfully.";
        public static string UpdateMailContactGeneralProfile => "Mail contact properties updated successfully.";
    }
}
