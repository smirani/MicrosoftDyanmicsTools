using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Collections;
using System.Xml.Linq;
using static System.IdentityModel.Tokens.SecurityTokenHandlerCollectionManager;

namespace ProjectImports
{
    internal class Program
    {
        public static EntityReference teamRoleProjectManager = new EntityReference("plus_teamrole" , new Guid("6e3c2bfb-07f5-e911-a960-000d3a34ebb5"));
        public static EntityReference teamRoleCOP = new EntityReference("plus_teamrole" , new Guid("c7315cac-e041-e911-a955-000d3a4f1a4f"));
        public static EntityReference clientRoleContractRepresentative = new EntityReference("plus_clientrole", new Guid("5e1cbe5f-48c4-e911-a85a-000d3a37cc26"));
        public static void Bootstrap()
        {
            var path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "log4net.config");
            Log<Program>.Configure(path);
        }
        static void Main(string[] args)
        {
            Bootstrap();

            // ************* MUST DO THESE STEPS BEFORE RUNNING LoadProjectsFromExcel ********
            // 
            // step 1. ensure the correct environment is set in App.config before running
            // step 2. A process must be turned off to prevent field deletion "DAI - On Project Create Map Over Key Fields"
            // step 3. ensure  filename is set to correct file.
            // *******************************************************************************
            // LoadProjectsFromExcel();

            // Update Opportunity DAI Connect Link
            UpdateOpportunitiesFromExcel();
            
        }
        public static void LoadProjectsFromExcel()
        {
            Log<Program>.Logger.Info("Project Imports Started");

            List<Project> projects = LoadFromExcel();

            Log<Program>.Logger.Info("Projects loaded");

            CreateNewProjects(projects);

            Log<Program>.Logger.Info("Project Imports Completed");
        }
        public static void UpdateOpportunitiesFromExcel()
        {
            Log<Program>.Logger.Info("Update Opportunities Started");

            List<Opportunity> opportunities = LoadOpportunitiesFromExcel();

            Log<Program>.Logger.Info("Opportunities loaded");

            UpdateOpportunities(opportunities);

            Log<Program>.Logger.Info("Update Opportunities Completed");
        }

        static List<Project> LoadFromExcel()
        {
            string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Files\CRM Project Upload Template (version 1)_MSP.xlsx";
            Application excelApp = new Application();
            Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
            Worksheet excelWorksheet = excelWorkbook.Sheets[1]; // assuming data is in first sheet
            Range excelRange = excelWorksheet.UsedRange;
            Log<Program>.Logger.Info(string.Format("{0} - {1} rows.", excelWorksheet.Name, excelRange.Rows.Count));
            List<Project> projects = new List<Project>();

            for (int row = 2; row <= excelRange.Rows.Count; row++)
            {
                if (excelRange.Cells[row, 1] == null || ((Range)excelRange.Cells[row, 1]).Value2 == null)
                {
                    //no more data, next sheet
                    row = excelRange.Rows.Count;
                    continue;
                }
                Double amount = 0;
                if ((((Range)excelRange.Cells[row, 13]).Value2).GetType() == typeof(double))
                {
                   amount = ((Range)excelRange.Cells[row, 13]).Value2;
                }

                Project project = new Project()
                {

                    ParentOracleProjectNumber = ((Range)excelRange.Cells[row, 1]).Value2.ToString(),
                    ParentProjectName = ((Range)excelRange.Cells[row, 2]).Value2,
                    Client = ((Range)excelRange.Cells[row, 3]).Value2,
                    OracleProjectNumber = ((Range)excelRange.Cells[row, 4]).Value2.ToString(),
                    OracleTaskNumber = ((Range)excelRange.Cells[row, 5]).Value2.ToString(),
                    Name = ((Range)excelRange.Cells[row, 6]).Value2,
                    OperatingUnit = ((Range)excelRange.Cells[row, 7]).Value2,
                    Description = ((Range)excelRange.Cells[row, 8]).Value2,
                    Sector = ((Range)excelRange.Cells[row, 9]).Value2,
                    SubSector = ((Range)excelRange.Cells[row, 10]).Value2,
                    Geography = (((Range)excelRange.Cells[row, 11]).Value2).Split(new[] { ", " }, StringSplitOptions.None),
                    ObligatedAmount = Convert.ToDecimal(amount),
                    StartDate = DateTime.FromOADate(((Range)excelRange.Cells[row, 14]).Value2).ToUniversalTime(),
                    EndDate = DateTime.FromOADate(((Range)excelRange.Cells[row, 15]).Value2).ToUniversalTime(),
                    HomeOfficePOC = ((Range)excelRange.Cells[row, 16]).Value2,
                    ProjectStaffPOC = ((Range)excelRange.Cells[row, 17]).Value2,
                    ClientPOCEmail = ((Range)excelRange.Cells[row, 18]).Value2,
                    ClientPOCFirstName = ((Range)excelRange.Cells[row, 19]).Value2,
                    ClientPOCLastName = ((Range)excelRange.Cells[row, 20]).Value2,
                    ClientPOCJobTitle = ((Range)excelRange.Cells[row, 21]).Value2,
                    DocLink = ((Range)excelRange.Cells[row, 22]).Value2,
                    Currency = "USD"
                };
                Log<Program>.Logger.Info(string.Format("{0} - {1}", row, project.Name));
                projects.Add(project);
            }
            excelWorkbook.Close(false);
            excelApp.Quit();

            return projects;
        }
        static void CreateNewProjects(List<Project> projects)
        {
            //use dictionaries to decrease runtime when uploading from file
            //dictionary for currencies, clients, sectors, countries, users, Buy-In Client Operating Units projects, contacts, projects

            Dictionary<string, Entity> allCurrencies = new Dictionary<string, Entity>();
            Dictionary<string, Entity> allClients = new Dictionary<string, Entity>();
            Dictionary<string, Entity> allSectors = new Dictionary<string, Entity>();
            Dictionary<string, Entity> allCountries = new Dictionary<string, Entity>();
            Dictionary<string, Entity> allUsers = new Dictionary<string, Entity>();
            Dictionary<string, Entity> allOperatingUnits = new Dictionary<string , Entity>();
            Dictionary<string, Entity> allContacts = new Dictionary<string, Entity>();
            Dictionary<string, Entity> allProjects = new Dictionary<string, Entity>();
            Dictionary<string, string> translatedClient = new Dictionary<string, string>
            {
                { "Bureau for Resilience and Food Security", "USAID Bureau for Resilience and Food Security" }
            };
            Dictionary<string, string> translatedSector = new Dictionary<string, string>
                {
                    { "Environment", "Environment/Resilience" }
                };

            foreach (Project project in projects)
            {
                int i = projects.IndexOf(project);
                // find matching isocurrencycode and set for entityReference
                AddReferenceData(allCurrencies, project.Currency, i, DynamicsHelper.GetCurrencyByIsoCode);
                Entity currencyEntity = allCurrencies[project.Currency];
                
                // find matching client by name and set for entityReference
                if (translatedClient.ContainsKey(project.Client)) project.Client = translatedClient[project.Client];
                AddReferenceData(allClients, project.Client, i, DynamicsHelper.GetClientsByName);
                Entity clientEntity = allClients[project.Client];

                // find matching sector by name and set for entityReference
                if (translatedSector.ContainsKey(project.Sector)) project.Sector = translatedSector[project.Sector];
                AddReferenceData(allSectors, project.Sector, i, DynamicsHelper.GetSectorByName);
                Entity sectorEntity = allSectors[project.Sector];

                // find matching practice area and set for entityReference
                Entity practiceArea = null;
                if (project.SubSector != null) practiceArea = AddPracticeArea(sectorEntity.Id, project.SubSector, i);

                // find matching country by name and set for entityReference
                EntityReferenceCollection countryRecordReferences = AddCountriesData(allCountries, project.Geography, i);
                Entity primaryCountry = allCountries[project.Geography[0]];

                // find matching Buy-In Client Operating Unit and set for entityReference
                AddReferenceData(allOperatingUnits, project.OperatingUnit, i, DynamicsHelper.GetBuyInClientOperatingUnit);
                if (!allOperatingUnits.Keys.Contains(project.OperatingUnit)) CreateOperatingUnit(allOperatingUnits, project.OperatingUnit, i);
                Entity OperatingUnitEntity = allOperatingUnits[project.OperatingUnit];

                // find HomeOfficePOC user by email and set for entityReference
                Entity UserEntity = null;
                if (project.HomeOfficePOC != null)
                {
                    AddReferenceData(allUsers, project.HomeOfficePOC, i, DynamicsHelper.GetUserByEmail);
                    UserEntity = allUsers[project.HomeOfficePOC];
                }

                // find ProjectStaffPOC user by email and set for entityReference
                Entity UserEntity2 = null;
                if (project.ProjectStaffPOC != null)
                {
                    AddReferenceData(allUsers, project.ProjectStaffPOC, i, DynamicsHelper.GetUserByEmail);
                    if (allUsers.Keys.Contains(project.ProjectStaffPOC)) UserEntity2 = allUsers[project.ProjectStaffPOC];
                }


                // find ClientPOC contact by email and set for entityReference
                Entity ContactEntity = null;
                if (project.ClientPOCEmail != null)
                {
                    AddReferenceData(allContacts, project.ClientPOCEmail, i, DynamicsHelper.GetContactByEmail);
                    if (!allContacts.Keys.Contains(project.ClientPOCEmail)) CreateContactRecord(allContacts, project.ClientPOCEmail, project.ClientPOCFirstName, project.ClientPOCLastName, project.ClientPOCJobTitle, i);
                    ContactEntity = allContacts[project.ClientPOCEmail];
                }
                

                // find parent project by parentProjectName and parentOraclePrjectNumber and set for entityReference (parent project, owner, and business unit)
                AddReferenceData(allProjects, project.ParentProjectName, i, project.ParentOracleProjectNumber, DynamicsHelper.GetProjectByNameAndOracleId);
                Entity ParentProjectEntity = allProjects[project.ParentProjectName];

                // set sub project number with project number - oracle task number
                string subProjectNumber = project.OracleProjectNumber.Substring(0, 7) + "-" + project.OracleTaskNumber.Substring(0, 3);

                
                // check for duplicate projects and create record
                if (DynamicsHelper.GetProjectByNameAndOracleId(project.Name, subProjectNumber) == null)
                {
                    Entity newRecord = new Entity("plus_project");
                    newRecord["plus_name"] = project.Name;
                    newRecord["plus_businessunitid"] = ParentProjectEntity.GetAttributeValue<EntityReference>("plus_businessunitid");
                    newRecord["plus_clientid"] = new EntityReference("account", clientEntity.Id);
                    newRecord["plus_oracleprojectnumber"] = subProjectNumber;
                    newRecord["plus_proposalid"] = ParentProjectEntity.GetAttributeValue<EntityReference>("plus_proposalid");
                    newRecord["dai_buyinclientoperatingunit"] = new EntityReference("dai_buyinclientoperatingunit", OperatingUnitEntity.Id);
                    newRecord["plus_projectdescription"] = project.Description;
                    newRecord["plus_startdate"] = project.StartDate;
                    newRecord["plus_enddate"] = project.EndDate;
                    if (project.ObligatedAmount != 0) newRecord["plus_obligatedamount"] = new Money(project.ObligatedAmount);
                    newRecord["transactioncurrencyid"] = new EntityReference("transactioncurrency", currencyEntity.Id);
                    newRecord["plus_daisectorid"] = new EntityReference("plus_daisector", sectorEntity.Id);
                    if (practiceArea != null) newRecord["dai_practicearea"] = new EntityReference("dai_practicearea", practiceArea.Id);
                    newRecord["plus_countryid"] = new EntityReference("plus_country", primaryCountry.Id);
                    newRecord["plus_parentprojectid"] = new EntityReference("plus_project", ParentProjectEntity.Id);
                    newRecord["ownerid"] = ParentProjectEntity.GetAttributeValue<EntityReference>("ownerid");
                    Guid recordId = DynamicsHelper.Create(newRecord);
                    Log<Program>.Logger.InfoFormat("Created new project with ID: {0}", recordId.ToString());
                    // add countries to the subgrid Other Countries
                    DynamicsHelper.AddCountriesToProject(recordId, countryRecordReferences);
                    // create and add DAI Team Member HomeOfficePOC
                    if (UserEntity != null) AddTeamMember(recordId, UserEntity.Id, teamRoleProjectManager, i);
                    // create and add DAI Team Member projectStaffPOC
                    if (UserEntity2 != null) AddTeamMember(recordId, UserEntity2.Id, teamRoleCOP, i);
                    // create and add client contact clientPOC
                    if (ContactEntity != null) AddContactClient(recordId, ContactEntity.Id, clientRoleContractRepresentative, i);
                }
                else
                {
                    Log<Program>.Logger.InfoFormat(" duplicate detected for {0} - {1} - {2}", (i + 1).ToString("0000#"), project.Name, subProjectNumber);
                    // append line to output if duplicate is detected
                }
            }
        }
        static void AddReferenceData(Dictionary<string, Entity> dictionary, string key, int index, string param, Func<string, string, Entity> method)
        {
            try
            {
                if (!dictionary.Keys.Contains(key))
                {
                    Entity entity = method(key, param);
                    if (entity != null) dictionary.Add(key, entity);
                    else throw new Exception("entity not found");
                }
            }
            catch (Exception e)
            {
                Log<Program>.Logger.InfoFormat("{0} - invalid {1}", (index + 1).ToString("0000#"), key);
            }
        }
        static void AddReferenceData(Dictionary<string, Entity> dictionary, string key, int index, Func<string, Entity> method)
        {
            try
            {
                if (!dictionary.Keys.Contains(key))
                {
                    Entity entity = method(key);
                    if (entity != null) dictionary.Add(key, entity);
                    else throw new Exception("entity not found");
                }
            }
            catch (Exception e)
            {
                Log<Program>.Logger.InfoFormat("{0} - invalid {1}",(index + 1).ToString("0000#"), key);
            }
        }

        static EntityReferenceCollection AddCountriesData(Dictionary<string, Entity> dictionary, string[] geography, int index)
        {
            EntityReferenceCollection countryRecordReferences = new EntityReferenceCollection();
            for (int n = 0; n < geography.Length; n++)
            {
                AddReferenceData(dictionary, geography[n], index, DynamicsHelper.FindCountryByName);
                Entity countryEntity = dictionary[geography[n]];
                countryRecordReferences.Add(countryEntity.ToEntityReference());
            }
            return countryRecordReferences;
        }
        static void AddTeamMember(Guid recordId, Guid userGuid,EntityReference teamRole, int index)
        {
            try
            {
                Entity newDaiTeamMemberRecord = new Entity("plus_daiteam");
                newDaiTeamMemberRecord["plus_memberid"] = new EntityReference("systemuser", userGuid);
                newDaiTeamMemberRecord["plus_roleid"] = teamRole;
                Guid daiTeamMemberRecordGuid = DynamicsHelper.Create(newDaiTeamMemberRecord);
                
                if (daiTeamMemberRecordGuid != Guid.Empty)
                {
                    Log<Program>.Logger.InfoFormat("Created new team member with ID: {0}", daiTeamMemberRecordGuid.ToString());
                    DynamicsHelper.AddUserToDAITeamForProject(recordId, daiTeamMemberRecordGuid);
                }
                else throw new Exception("team Member not created");
            }
            catch (Exception e)
            {
                Log<Program>.Logger.InfoFormat("{0} - team member not created {1}", (index + 1).ToString("0000#"), userGuid);
            }       
        }   
        static void AddContactClient(Guid recordId, Guid contactGuid, EntityReference clientRole, int index)
        {
            try
            {
                Entity newClientContactRecord = new Entity("plus_clientcontact");
                newClientContactRecord["plus_contact"] = new EntityReference("contact", contactGuid);
                newClientContactRecord["plus_clientrole"] = clientRole;
                Guid clientContactGuid = DynamicsHelper.Create(newClientContactRecord);
                if (clientContactGuid != Guid.Empty)
                {
                    Log<Program>.Logger.InfoFormat("Created new Client Contact with ID: {0}", clientContactGuid.ToString());
                    DynamicsHelper.AddClientToProject(recordId, clientContactGuid);
                }
                else throw new Exception("client contact not created");
            }
            catch (Exception e)
            {
                Log<Program>.Logger.InfoFormat("{0} - client contact not created {1}", (index + 1).ToString("0000#"), contactGuid);
            }
        }
        static Entity AddPracticeArea(Guid sectorGuid, string subSector, int index)
        {
            Entity practiceArea = DynamicsHelper.GetPracticeAreaByName(sectorGuid, subSector);
            if (practiceArea == null)
            {
                Log<Program>.Logger.InfoFormat("{0} - Practice area not found {1}", (index + 1).ToString("0000#"), subSector);
                throw new Exception("Practice area not found");
            }
            return practiceArea;
        }
        static void CreateOperatingUnit(Dictionary<string, Entity> dictionary, string name, int index)
        {
            Entity newOperatingUnit = new Entity("dai_buyinclientoperatingunit");
            newOperatingUnit["dai_name"] = name;
            Guid newOperatingUnitGuid = DynamicsHelper.Create(newOperatingUnit);
            Entity newOperatingUnitEntity;
            if (newOperatingUnitGuid != Guid.Empty)
            {
                Log<Program>.Logger.InfoFormat("Created new Operating Unit with ID: {0}", newOperatingUnitGuid.ToString());
                newOperatingUnitEntity = DynamicsHelper.Get("dai_buyinclientoperatingunit", newOperatingUnitGuid);
                dictionary.Add(name, newOperatingUnitEntity);
            }
            else
            {
                Log<Program>.Logger.InfoFormat("{0} - Operating Unit not created {1}", (index + 1).ToString("0000#"), name);
                throw new Exception("Operating Unit not created");
            }
        }
        static void CreateContactRecord(Dictionary<string, Entity> dictionary, string email, string firstName, string lastName, string jobTitle, int index)
        {
            Entity newContact = new Entity("contact");
            newContact["emailaddress1"] = email;
            newContact["firstname"] = firstName;
            newContact["lastname"] = lastName;
            newContact["jobtitle"] = jobTitle;
            Guid newContactGuid = DynamicsHelper.Create(newContact);
            Entity newContactEntity;
            if (newContactGuid != Guid.Empty)
            {
                Log<Program>.Logger.InfoFormat("Created new Contact with ID: {0}", newContactGuid.ToString());
                newContactEntity = DynamicsHelper.Get("contact", newContactGuid);
                dictionary.Add(email, newContactEntity);
            }
            else
            {
                Log<Program>.Logger.InfoFormat("{0} - Contact not created {1}", (index + 1).ToString("0000#"), email);
                throw new Exception("Contact not created");
            }
        }
        
        static List<Opportunity> LoadOpportunitiesFromExcel()
        {
            string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Files\DAI HMG - BDU June 2023 Missing Links.xlsx";
            Application excelApp = new Application();
            Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
            Worksheet excelWorksheet = excelWorkbook.Sheets[1]; // assuming data is in first sheet
            Range excelRange = excelWorksheet.UsedRange;
            Log<Program>.Logger.Info(string.Format("{0} - {1} rows.", excelWorksheet.Name, excelRange.Rows.Count));
            List<Opportunity> opportunities = new List<Opportunity>();
            for (int row = 2; row <= excelRange.Rows.Count; row++)
            {
                if (excelRange.Cells[row, 1] == null || ((Range)excelRange.Cells[row, 1]).Value2 == null)
                {
                    //no more data, next sheet
                    row = excelRange.Rows.Count;
                    continue;
                }
                Opportunity opportunity = new Opportunity()
                {
                    OpportunityId = ((Range)excelRange.Cells[row, 1]).Value2,
                    Name = ((Range)excelRange.Cells[row, 6]).Value2,
                    DAIConnectLink = ((Range)excelRange.Cells[row, 17]).Value2
                };
                Log<Program>.Logger.Info(string.Format("{0} - {1}", row, opportunity.Name));
                opportunities.Add(opportunity);
            }
            excelWorkbook.Close(false);
            excelApp.Quit();

            return opportunities;
        }
        static void UpdateOpportunities(List<Opportunity> opportunities)
        {
            foreach(Opportunity opportunity in opportunities)
            {
                try
                {
                    Entity OpportunityRecord = DynamicsHelper.Get("opportunity", Guid.Parse(opportunity.OpportunityId));
                    OpportunityRecord["dai_daiconnectlink"] = opportunity.DAIConnectLink;
                    DynamicsHelper.Update(OpportunityRecord);
                    Log<Program>.Logger.InfoFormat("{0} - updated url link for {1}", opportunities.IndexOf(opportunity), opportunity.Name);

                }
                catch (Exception e)
                {
                    Log<Program>.Logger.InfoFormat("{0} - could not update url link for {1}", opportunities.IndexOf(opportunity), opportunity.Name);
                }
                
            }
        }
    }
}
