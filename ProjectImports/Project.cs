using System;

namespace ProjectImports
{
    public class Project
    {
        // add attributes for project
        public string Name { get; set; }
        public string OracleProjectNumber { get; set; }
        public string OracleTaskNumber { get; set; }
        public string ParentProjectName { get; set; }
        public string ParentOracleProjectNumber { get; set; }
        public string OperatingUnit { get; set; }
        public string Client { get; set; }
        public string Description { get; set; }
        public string Sector { get; set; }
        public string SubSector { get; set; }
        public string[] Geography { get; set; }
        public string Currency { get; set; }
        public decimal ObligatedAmount { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string HomeOfficePOC { get; set; }
        public string ProjectStaffPOC { get; set; }
        public string ClientPOCEmail { get; set; }
        public string ClientPOCFirstName { get; set; }
        public string ClientPOCLastName { get; set; }
        public string ClientPOCJobTitle { get; set; }
        public string DocLink { get; set; }
        
    }
}
