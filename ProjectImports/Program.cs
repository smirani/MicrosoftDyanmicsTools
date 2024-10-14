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
using Microsoft.Crm.Sdk.Messages;

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
            //LoadProjectsFromExcel();

            // Update Opportunity DAI Connect Link
            // UpdateOpportunitiesFromExcel();

            // Remove Duplicate Questions from Proposals
            // RemoveDuplicateQuestionsFromProposals();

            // Delete Unmanaged Solutions
            RemoveUnmanagedSolutionsFromEnvironment();

            // Update Custom Question Responses
            // UpdateAllCustomQuestionResponses();
            
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
        public static void RemoveDuplicateQuestionsFromProposals()
        {
            Log<Program>.Logger.Info("Remove Duplicate Questions Started");
            DeleteDuplicateCustomQuestions();
            Log<Program>.Logger.Info("Remove Duplicate Questions Completed");
        }
        public static void RemoveUnmanagedSolutionsFromEnvironment()
        {
            Log<Program>.Logger.Info("Remove Unmanaged Solutions Started");
            DeleteUnmanagedSolutions();
            Log<Program>.Logger.Info("Remove Unmanaged Solutions Completed");
        }

        public static void UpdateAllCustomQuestionResponses()
        {
            Log<Program>.Logger.Info("Update Custom Question Responses Started");
            UpdateResponseCustomQuestionResponses();
            Log<Program>.Logger.Info("Update Custom Question Responses Completed");
        }

        static List<Project> LoadFromExcel()
        {
            string filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Files\CRM Project Upload Template (version 1)_MSP(2).xlsx";
            Application excelApp = new Application();
            Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
            Worksheet excelWorksheet = excelWorkbook.Sheets[7]; // 1, assuming data is in first sheet
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
        static void DeleteDuplicateCustomQuestions()
        {
            // get all proposals that have a solicitation
            string query = 
                    "<fetch>" +
                        "<entity name='cdr_contractapplication'>" +
                            "<attribute name='cdr_contractapplicationid' />" +
                            "<attribute name='cdr_name' />" +
                            "<filter type='and'>" +
                                "<condition attribute='dai_solicitationid' operator='not-null' />" +
                                "<condition attribute='statecode' operator='eq' value='0'/>" +
                            "</filter>" +
                        "</entity>" +
                    "</fetch>";
            EntityCollection proposals = DynamicsHelper.FetchXmlQuery(query);

            Log<Program>.Logger.InfoFormat("Found {0} proposals with solicitations", proposals.Entities.Count);
            
            // for each proposal, get all custom questions
            foreach (Entity proposal in proposals.Entities)
            {
                // get all custom questions for proposal
                string query2 =
                        "<fetch>" +
                            "<entity name='gm_grantperformancecriterion'>" +
                                "<attribute name='gm_name' />" +
                                "<attribute name='gm_metrictype' />" +
                                "<attribute name='gm_displayorder' />" +
                                "<attribute name='cdr_yesnorespactual' />" +
                                "<attribute name='cdr_wholerespactual' />" +
                                "<attribute name='cdr_commentactual' />" +
                                "<attribute name='cdr_applicationcontract' />" +
                                "<attribute name='isi_concept' />" +
                                "<attribute name='isi_optionsetresponse' />" +
                                "<attribute name='gm_floatnumberresponse' />" +
                                "<attribute name='gm_grantperformancecriterionid' />" +
                                "<attribute name= 'isi_answered' />" +
                                "<filter type='and'>" +
                                "<condition attribute='statecode' operator='eq' value='0' />" +
                                "</filter>" +
                                "<link-entity name='cdr_contractapplication' from='cdr_contractapplicationid' to='cdr_applicationcontract' link-type='inner' alias='ac'>" +
                                "<filter type='and'>" +
                                    "<condition attribute='cdr_contractapplicationid' operator='eq' uitype='cdr_contractapplication' value='" + proposal.GetAttributeValue<Guid>("cdr_contractapplicationid").ToString() + "' />" +
                                "</filter>" +
                                "</link-entity>" +
                            "</entity>" +
                        "</fetch>";
                EntityCollection customQuestions = DynamicsHelper.FetchXmlQuery(query2);
                Log<Program>.Logger.InfoFormat("Found {0} custom questions for proposal {1}", customQuestions.Entities.Count, proposal.GetAttributeValue<string>("cdr_name"));
                // for each custom question, check for duplicates and delete
                foreach (Entity customQuestion in customQuestions.Entities)
                {
                    //check for duplicates
                    List<Entity> duplicates = customQuestions.Entities.Where(x => x.GetAttributeValue<string>("gm_name") == customQuestion.GetAttributeValue<string>("gm_name")).ToList();
                    //if duplicates exist, log them
                    
                    if (duplicates.Count > 1)
                    {
                        Log<Program>.Logger.InfoFormat("Found {0} duplicates for custom question {1} in proposal {2}", duplicates.Count, customQuestion.GetAttributeValue<string>("gm_name"), proposal.GetAttributeValue<string>("cdr_name"));
                        //delete duplicate question where answered = false
                        for (int i = 1; i < duplicates.Count; i++)
                        {
                            if (duplicates[i].GetAttributeValue<bool>("isi_answered") == false)
                            {
                                string duplicateQuestion = duplicates[i].GetAttributeValue<string>("gm_name");
                                

                                try
                                {
                                    DynamicsHelper.Delete("gm_grantperformancecriterion", duplicates[i].Id);
                                }
                                catch (System.ServiceModel.FaultException e)
                                {
                                    Log<Program>.Logger.InfoFormat("Error deleting duplicate custom question {0} in proposal {1} duplicates count = {2}", duplicateQuestion, proposal.GetAttributeValue<string>("cdr_name"), duplicates.Count);
                                    Log<Program>.Logger.Info(e.Message);
                                }
                                catch (Exception ex)
                                {
                                   
                                    throw;
                                }
                                
                                Log<Program>.Logger.InfoFormat("Deleted duplicate custom question {0} in proposal {1} duplicates count = {2}", duplicateQuestion, proposal.GetAttributeValue<string>("cdr_name"), duplicates.Count) ;
                            }
                        }
                    }

                }
            }  
        }
        static void DeleteUnmanagedSolutions()
        {
            // get all unmanaged solutions
            string query = 
                    "<fetch>" +
                        "<entity name='solution'>" +
                            "<attribute name='friendlyname' />" +
                            "<attribute name='version' />" +
                            "<attribute name='solutionid' />" +
                            "attribute name='uniquename' />" +
                            "<filter type='and'>" +
                                "<condition attribute='ismanaged' operator='eq' value='0' />" +
                                "<condition attribute='isvisible' operator='eq' value='1' />" +
                                "<condition attribute='uniquename' operator='not-like' value='Default%' />" +
                                //"<condition attribute='modifiedon' operator='lt' value='" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd") + "' />" +
                                "<condition attribute='uniquename' operator='not-like' value='WO%' />" +
                            "</filter>" +
                        "</entity>" +
                    "</fetch>";
            EntityCollection solutions = DynamicsHelper.FetchXmlQuery(query);

            Log<Program>.Logger.InfoFormat("Found {0} unmanaged solutions", solutions.Entities.Count);
            
            // for each solution, delete
            foreach (Entity solution in solutions.Entities)
            {
                //user input to confirm deletion (comment out to run without confirmation)
                // Log<Program>.Logger.InfoFormat("Delete unmanaged solution {0} (y/n)?", solution.GetAttributeValue<string>("friendlyname"));
                // string input = Console.ReadLine();
                // if (input != "y") continue;

                try
                {
                    DynamicsHelper.Delete("solution", solution.Id);
                    Log<Program>.Logger.InfoFormat("Deleted unmanaged solution {0}", solution.GetAttributeValue<string>("friendlyname"));
                }
                catch (System.ServiceModel.FaultException e)
                {
                    Log<Program>.Logger.InfoFormat("Error deleting unmanaged solution {0}", solution.GetAttributeValue<string>("friendlyname"));
                    Log<Program>.Logger.Info(e.Message);
                }
                catch (Exception ex)
                {
                    throw;
                }
            }
        }
        static void UpdateResponseCustomQuestionResponses()
        {
            // retrieve all custom questions
            string query = 
                    "<fetch>" +
                        "<entity name='gm_grantperformancecriterion'>" +
                            "<attribute name='gm_name' />" +
                            "<attribute name='gm_metrictype' />" +
                            "<attribute name='gm_displayorder' />" +
                            "<attribute name='cdr_yesnorespactual' />" +
                            "<attribute name='cdr_wholerespactual' />" +
                            "<attribute name='cdr_commentactual' />" +
                            "<attribute name='cdr_applicationcontract' />" +
                            "<attribute name='isi_concept' />" +
                            "<attribute name='isi_optionsetresponse' />" +
                            "<attribute name='gm_floatnumberresponse' />" +
                            "<attribute name='gm_grantperformancecriterionid' />" +
                            "<attribute name='isi_percentageresponse' />" +
                            "<attribute name= 'isi_answered' />" +
                            "<filter type='and'>" +
                            "<condition attribute='statecode' operator='eq' value='0' />" +
                            "<condition attribute='isi_answered' operator='eq' value='true' />" +
                            "<condition attribute='gm_metrictype' operator='eq' value='930660002' />" +
                            "</filter>" +
                        "</entity>" +
                    "</fetch>";
            EntityCollection customQuestions = DynamicsHelper.FetchXmlQuery(query);
            Log<Program>.Logger.InfoFormat("Found {0} custom questions", customQuestions.Entities.Count);
            // for each custom question check metrictype and update response
            foreach (Entity customQuestion in customQuestions.Entities)
            {

                string newResponse = "";
                // switch on gm_metrictype
                switch (customQuestion.GetAttributeValue<OptionSetValue>("gm_metrictype").Value)
                {
                    case 100000000: // Text Only
                        if (customQuestion.GetAttributeValue<string>("cdr_commentactual") != null)
                        {
                            newResponse = customQuestion.GetAttributeValue<string>("cdr_commentactual");
                            // if newResponse is more than 4000 characters, truncate
                            if (newResponse.Length > 4000)
                            {
                                newResponse = newResponse.Substring(0, 4000);
                            }
                        }
                        break;
                    case 605730000: // Option Set Single
                        if (customQuestion.GetAttributeValue<EntityReference>("isi_optionsetresponse") != null)
                        {
                            newResponse = customQuestion.GetAttributeValue<EntityReference>("isi_optionsetresponse").Name;
                        }
                        break;
                    case 930660002: // Yes/No 
                        if (customQuestion.GetAttributeValue<bool>("cdr_yesnorespactual").Equals(true))
                        {
                            newResponse = "Yes";
                        }
                        else
                        {
                            newResponse = "No";
                        }
                        break;
                    case 930660000: // Whole Number
                            newResponse = customQuestion.GetAttributeValue<int>("cdr_wholerespactual").ToString();
                        break;
                    case 930660001: // Decimal Number
                        
                        if (customQuestion.FormattedValues.ContainsKey("gm_floatnumberresponse"))
                        {
                            newResponse = customQuestion.FormattedValues["gm_floatnumberresponse"];
                        }
                          
                            
                        break;
                    case 605730001: // Percentage
                            newResponse = customQuestion.GetAttributeValue<int>("isi_percentageresponse").ToString();
                        break;
                    case 930660004: // Country single select
                        if (customQuestion.GetAttributeValue<EntityReference>("cdr_countryresponse") != null)
                        {
                            newResponse = customQuestion.GetAttributeValue<EntityReference>("cdr_countryresponse").Name;
                        }
                        break;
                    case 930660005: // Country multi select
                        // get all countries related to the custom question
                       string query2 = 
                            "<fetch>" +
                                "<entity name='plus_country'>" +
                                    "<attribute name='plus_name' />" +
                                    "<attribute name='plus_countryid' />" +
                                    "<filter type='and'>" +
                                        "<condition attribute='statecode' operator='eq' value='0' />" +
                                    "</filter>" +
                                    "<link-entity name='dai_gm_grantperformancecriterion_plus_count' from='plus_countryid' to='plus_countryid' visible='false' intersect='true'>" +
                                        "<link-entity name='gm_grantperformancecriterion' from='gm_grantperformancecriterionid' to='gm_grantperformancecriterionid' alias='ab'>" +
                                            "<filter type='and'>" +
                                                "<condition attribute='gm_grantperformancecriterionid' operator='eq' uitype='gm_grantperformancecriterion' value='" + customQuestion.Id.ToString() + "' />" +
                                            "</filter>" +
                                        "</link-entity>" +
                                    "</link-entity>" +
                                "</entity>" +
                            "</fetch>"; 
                        EntityCollection countries = DynamicsHelper.FetchXmlQuery(query2);
                        // build a comma separated string of country names
                        foreach (Entity country in countries.Entities)
                        {
                            newResponse += country.GetAttributeValue<string>("plus_name") + ", ";
                        }
                        newResponse = newResponse.TrimEnd(',', ' ');
                        break;
                    case 930660003: // Option Set Multiple
                        // get all metric option set responses related to the custom question
                        string query3 = 
                            "<fetch>" +
                                "<entity name='isi_metricoptionsetanswer'>" +
                                    "<attribute name='isi_metricoptionsetanswerid' />" +
                                    "<attribute name='isi_name' />" +
                                    "<link-entity name='dai_gm_grantperformancecriterion_isi_metric' from='isi_metricoptionsetanswerid' to='isi_metricoptionsetanswerid' visible='false' intersect='true'>" +
                                        "<link-entity name='gm_grantperformancecriterion' from='gm_grantperformancecriterionid' to='gm_grantperformancecriterionid' alias='af'>" +
                                            "<filter type='and'>" +
                                                "<condition attribute='gm_grantperformancecriterionid' operator='eq' uitype='gm_grantperformancecriterion' value='" + customQuestion.Id.ToString() + "' />" +
                                            "</filter>" +
                                        "</link-entity>" +
                                    "</link-entity>" +
                                "</entity>" +
                            "</fetch>";
                        EntityCollection responses = DynamicsHelper.FetchXmlQuery(query3);
                        // build a comma separated string of response names 
                        foreach (Entity response in responses.Entities)
                        {
                            newResponse += response.GetAttributeValue<string>("isi_name") + ", ";
                        }
                        newResponse = newResponse.TrimEnd(',', ' ');
                        break;
                    default:
                        break;
                }
                if (newResponse != "")
                {
                    try
                    {
                        customQuestion["dai_response"] = newResponse;
                        DynamicsHelper.Update(customQuestion);
                        Log<Program>.Logger.InfoFormat("Updated response for custom question {0} to {1}", customQuestion.GetAttributeValue<string>("gm_name"), newResponse);
                    }
                    catch ( Exception e)
                    {
                        var temp = e.Message;
                    }
                }
            }

        }
    }
}
