using System.Linq;

namespace TestProject1
{
    public class Record
    {
        private string currentAssignmentGroup;
        public string CurrentAssignmentGroup 
        {
            get { return currentAssignmentGroup; } 
            set {
                    currentAssignmentGroup = value;
                    if (Const.adessoAssignemntGroupKeyWords.Any(group => value.ToLower().Contains(group)))
                    {
                        LastAdessoAssignmentGroup = value;
                        if (String.IsNullOrEmpty(FirstAdessoAssignmentGroup))
                        {
                            FirstAdessoAssignmentGroup = value;
                        }
                    }
                } 
        }
        public string ColleagueName
        { 
            get; 
            set; 
        }
        public string LastAdessoSupporter { get; set; }
        public void AddColleagueToConnectedSupportersList()
        {
            LastAdessoSupporter = ColleagueName;
            if(!AllConnectedAdessoSupporters.Contains(ColleagueName))
            {
                AllConnectedAdessoSupporters += ColleagueName + ",";
            }
        }

        public void SetAssignmentGroupFlagsToTrueIfNecessary()
        {
            var group = CurrentAssignmentGroup.ToLower();
            if (group.Contains(Const.internalAppsStr))
            {
                this.RelatedWithInternalApplications = true;
            }
            else if (group.Contains(Const.bluenetStr))
            {
                this.RelatedWithBluenet = true;
            }
            else if (group.Contains(Const.ccptStr))
            {
                this.RelatedWithCCPT = true;
            }
            else if (group.Contains(Const.cloudMoveStr))
            {
                this.RelatedWithCloudMove = true;
            }
            else if (group.Contains(Const.ltsStr))
            {
                this.RelatedWithLTS = true;
            }
            else if (group.Contains(Const.capacityToolStr))
            {
                this.RelatedWithCapacityTool = true;
            }
            else if (group.Contains(Const.repairNonitorStr))
            {
                this.RelatedWithRepairMonitor = true;
            }

        }

        public string AllConnectedAdessoSupporters { get; set;}
        public string Application { get; set; }
        public string Service { get; set; }
        public string ServiceOffering { get; set; }
        public string Priority { get; set; }
        public string TicketType { get; set; }
        public string Status { get; set; }
        public string TicketOpenedAt { get; set; }
        public string TicketForwardedAt { get; set; }
        public string TicketClosedAt { get; set; }
        public string FirstAdessoAssignmentGroup { get; set; }
        public string LastAdessoAssignmentGroup { get; set; }
        public string AssignmentTime { get; set; }
        public string ReactionTime { get; set; }
        public string ResolutionTime { get; set; }
        public string AffectedApplication { get; set; }
        public string TicketId { get; set; }
        public string Description { get; set;}

        public bool RelatedWithInternalApplications { get; set; }
        public bool RelatedWithBluenet { get; set; }
        public bool RelatedWithCCPT { get; set; }
        public bool RelatedWithCapacityTool { get; set; }
        public bool RelatedWithLTS { get; set; }
        public bool RelatedWithRepairMonitor { get; set; }
        public bool RelatedWithCloudMove { get; set; }

        public int RelatedAssignmentGroupsCount {
            get
            {
                return Convert.ToInt32(this.RelatedWithBluenet) +
                   Convert.ToInt32(this.RelatedWithInternalApplications) +
                   Convert.ToInt32(this.RelatedWithCCPT) +
                   Convert.ToInt32(this.RelatedWithCloudMove) +
                   Convert.ToInt32(this.RelatedWithLTS) +
                   Convert.ToInt32(this.RelatedWithRepairMonitor) +
                   Convert.ToInt32(this.RelatedWithCapacityTool) 
                    ;
            } 
        }
    }
}