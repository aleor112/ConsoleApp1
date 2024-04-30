using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject1
{
    public static class Const
    {
        public static string urlPlaceholder = "";
        public static int BatchSizeTicketGenerate = 1000;
        public static int TicketsPerPage = 100;
        public static string AssignedTo = "AssignedTo";
        public static string AssignmentGroup = "AssignmentGroup";
        public static string AssignmentTime = "AssignmentTime";
        public static string PreviousAssignmentTime = "PreviousAssignmentTime";
        public static string Category = "Category";
        public static string Priority = "Priority";
        public static string TicketOpenedAt = "TicketOpenedAt";
        public static string TicketClosedAt = "TicketClosedAt";
        public static string ReactionTime = "ReactionTime";
        public static string ResolutionTime = "ResolutionTime";
        public static string PreviousReactionTime = "PreviousReactionTime";
        public static string PreviousAssignmentGroup = "PreviousAssignmentGroup";
        public static string TicketId = "TicketId";
        public static string Service = "Service";
        public static string ServiceOffering = "ServiceOffering";
        public static string Description = "Description";
        public static string ForwardedAt = "ForwardedAt";
        public static string Application = "Application";
        public static string Completed = "Completed";
        public static string State = "State";
        public static string HasAdessoGroup = "HasAdessoGroup";
        public static string HasCreatedRecord = "HasCreatedRecord";
        public static string internalAppsStr = "internalapp";
        public static string repairNonitorStr = "repairmonitor";
        public static string cloudMoveStr = "cloud move";
        public static string ltsStr = "lts";
        public static string bluenetStr = "bluenet";
        public static string capacityToolStr = "capacitytool";
        public static string ccptStr = "ccpt";
        public static int numberOfTabs = 10;
        public static List<string> applicationWords = new List<string>() { "onlineapplication","sales2go","cad library","cadlibrary","bluenet", 
            "speiseleitsystem","confservice","citavi", " cpo ", " mpm ", "machinelogbook", "qmsm","qmnotification","q3notification", "maintenance other" };
        public static List<string> adessoAssignemntGroupKeyWords = new List<string>() { "adesso", "bluenet" };

    }
}
