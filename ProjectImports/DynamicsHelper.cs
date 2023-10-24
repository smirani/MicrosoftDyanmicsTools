using Microsoft.Crm.Sdk.Messages;
using Microsoft.Identity.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.WebServiceClient;
using System;
using System.Configuration;
using System.Linq;

namespace ProjectImports
{
    public class DynamicsHelper
    {
        private static readonly string _dynamicsApi = ConfigurationManager.AppSettings["DynamicsApi"];
        private static readonly Uri _orgService = new Uri(_dynamicsApi + "/XRMServices/2011/Organization.svc/web?SdkClientVersion=9.0");
        private static IConfidentialClientApplication _clientApp = default;

        public enum OrgType
        {
            Implementer = 990260000,
            Client = 990260001,
            Grantee = 990260002,
            Vendor = 990260003,
            Other = 990260005
        };

        public enum DocumentSource
        {
            MDA = 605730000,
            Portal = 605730000
        }

        public enum DocumentStage
        {
            Concept = 605730002
        }

        public static EntityReference DocumentTypeConceptPlan => new EntityReference("isi_gmdocumenttype", new Guid("4165214B-E559-EC11-8F8F-0022480A6828"));
        public static EntityReference DocumentTypeOther => new EntityReference("isi_gmdocumenttype", new Guid("59157389-1973-EC11-8941-0022480B3952"));
        public static OptionSetValue StageConcept => new OptionSetValue((int)DocumentStage.Concept);

        private static string AccessToken2()
        {
            string clientID = "039ddc98-daf8-4c39-9ea0-27327a3419b2";
            string clientSecret = "d6qs3J6HOwlpuYDvtd4/Jw/sDBEt9lOoUjzVr6PIKsU=";
            string authority = "https://login.microsoftonline.com/7107113d-e20b-4c20-a4ce-553cabbf686d";
            string[] scopes = new string[] { _dynamicsApi + "/.default" };
            string token = default;

            try
            {
                if (_clientApp == default)
                {
                    _clientApp = ConfidentialClientApplicationBuilder.Create(clientID)
                                                                    .WithClientSecret(clientSecret)
                                                                    .WithAuthority(authority)
                                                                    .WithLegacyCacheCompatibility(false)
                                                                    .WithCacheOptions(CacheOptions.EnableSharedCacheOptions)
                                                                    .Build(); 
                }
                Microsoft.Identity.Client.AuthenticationResult result = _clientApp.AcquireTokenForClient(scopes).ExecuteAsync().Result;
                token = result.AccessToken;
            }
            catch (Exception e)
            {
                Log<Program>.Logger.Error(e);
            }

            return token;
        }

        public static EntityCollection FetchXmlQuery(string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return null;

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                return service.RetrieveMultiple(new FetchExpression(query));
            }
        }

        public static OrganizationResponse AddUsersToTeam(Entity[] users, Entity team)
        {
            return AddUsersToTeam(users.Select(entity => entity.Id).ToArray(), team.Id);
        }

        public static OrganizationResponse AddUsersToTeam(Guid[] userIds, Guid teamId)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                AddMembersTeamRequest addRequest = new AddMembersTeamRequest
                {
                    TeamId = teamId,
                    MemberIds = userIds
                };

                return service.Execute(addRequest);
            }
        }

        public static OrganizationResponse RemoveUsersFromTeam(Guid[] userIds, Entity team)
        {
            return RemoveUsersFromTeam(userIds, team.Id);
        }

        public static OrganizationResponse RemoveUsersFromTeam(Guid[] userIds, Guid teamId)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                RemoveMembersTeamRequest addRequest = new RemoveMembersTeamRequest
                {
                    TeamId = teamId,
                    MemberIds = userIds
                };

                return service.Execute(addRequest);
            }
        }

        public static void RemoveRolesFromUser(EntityReferenceCollection roles, Entity user)
        {
            RemoveRolesFromUser(roles, user.Id);
        }

        public static void RemoveRolesFromUser(EntityReferenceCollection roles, Guid userId)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                service.Disassociate("systemuser", userId, new Relationship("systemuserroles_association"), roles);
            }
        }

        public static void AddRolesToUser(Entity user, EntityReferenceCollection roles)
        {
            AddRolesToUser(user.Id, roles);
        }

        public static void AddRolesToUser(Guid userId, EntityReferenceCollection roles)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                service.Associate("systemuser", userId, new Relationship("systemuserroles_association"), roles);
            } 
        }

        public static Entity GetRoleByName(string name, Entity businessUnit)
        {
            return GetRoleByName(name, businessUnit.Id);
        }

        public static Entity GetUserByEmail(string email)
        {
            ConditionExpression condition = new ConditionExpression("systemuser", "internalemailaddress", ConditionOperator.Equal, email);
            QueryExpression query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet(true)
            };
            query.Criteria.AddCondition(condition);

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }

        public static Entity GetRoleByName(string name, Guid businessUnit)
        {
            ConditionExpression condition = new ConditionExpression("role", "name", ConditionOperator.Equal, name);
            ConditionExpression conditionBu = new ConditionExpression("role", "businessunitid", ConditionOperator.Equal, businessUnit);
            QueryExpression query = new QueryExpression("role");
            query.Criteria.AddCondition(condition);
            query.Criteria.AddCondition(conditionBu);

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }

       
            public static Entity GetTeamByName(string name)
        {
            ConditionExpression condition = new ConditionExpression("team", "name", ConditionOperator.Equal, name);
            QueryExpression query = new QueryExpression("team");
            query.Criteria.AddCondition(condition);

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }

        // <summary>
        // Gets the business unit by its name.
        // </summary>
        // <param name="name">The business unit name.</param>
        // <returns>Returns entity representing business unit or null.</returns>
        public static Entity GetBusinessUnitByName(string name)
        {
            ConditionExpression condition = new ConditionExpression("businessunit", "name", ConditionOperator.Equal, name);
            QueryExpression query = new QueryExpression("businessunit");
            query.Criteria.AddCondition(condition);

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }

        public static DataCollection<Entity> FindUserTeams(Entity user)
        {
            return FindUserTeams(user.Id);
        }
        public static DataCollection<Entity> FindUserTeams(Guid userId)
        {
            QueryExpression query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet(true)
            };
            LinkEntity link = query.AddLink("teammembership", "teamid", "teamid");
            link.LinkCriteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, userId));

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                return service.RetrieveMultiple(query).Entities;
            }
        }

        public static DataCollection<Entity> FindUserRoles(Entity user)
        {
            return FindUserRoles(user.Id);
        }

        public static DataCollection<Entity> FindUserRoles(Guid userId)
        {
            QueryExpression queryExpression = new QueryExpression
            {
                EntityName = "role",
                ColumnSet = new ColumnSet(true)
            };

            LinkEntity leRoleToUserRole = new LinkEntity
            {
                LinkFromEntityName = "role",
                LinkFromAttributeName = "roleid",
                LinkToEntityName = "systemuserroles",
                LinkToAttributeName = "roleid"
            };

            LinkEntity leUserRoleToUser = new LinkEntity
            {
                LinkFromEntityName = "systemuserroles",
                LinkFromAttributeName = "systemuserid",
                LinkToEntityName = "systemuser",
                LinkToAttributeName = "systemuserid"
            };

            ConditionExpression ce = new ConditionExpression
            {
                AttributeName = "systemuserid",
                Operator = ConditionOperator.Equal
            };
            ce.Values.Add(userId);

            leUserRoleToUser.LinkCriteria = new FilterExpression();
            leUserRoleToUser.LinkCriteria.AddCondition(ce);
            leRoleToUserRole.LinkEntities.Add(leUserRoleToUser);
            queryExpression.LinkEntities.Add(leRoleToUserRole);

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                return service.RetrieveMultiple(queryExpression).Entities;
            }
        }

        public static DataCollection<Entity> FindBpfByName(string name)
        {
            QueryExpression query = new QueryExpression("workflow")
            {
                ColumnSet = new ColumnSet(true)
            };
            query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, name));

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                return service.RetrieveMultiple(query).Entities;
            }
        }

        public static Entity GetProjectCsmById(Guid id)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                Entity entity = service.Retrieve("plus_projectcsm", id, new ColumnSet(true));
                return entity;
            }
        }

        public static Entity WhoAmI()
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                WhoAmIResponse whoAmIResponse = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                Entity entity = service.Retrieve("systemuser", whoAmIResponse.UserId, new ColumnSet(true));
                return entity;
            }
        }

        public static void MakeAdmin(string email)
        {
            Entity user = GetUserByEmail(email);

            string dest = ConfigurationManager.AppSettings["DynamicsDestination"];
            string bu = dest.Substring(8, dest.IndexOf(".") - 8);
            Entity businessUnit = GetBusinessUnitByName(bu);

            if (user.GetAttributeValue<EntityReference>("businessunitid").Id != businessUnit.Id)
            {
                SetBusinessSystemUserRequest bsRequest = new SetBusinessSystemUserRequest
                {
                    BusinessId = businessUnit.Id,
                    UserId = user.Id,
                    ReassignPrincipal = user.ToEntityReference()
                };

                using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
                {
                    SetBusinessSystemUserResponse bsResponse = (SetBusinessSystemUserResponse)service.Execute(bsRequest);
                }
            }

            DataCollection<Entity> userRoles = FindUserRoles(user.Id);
            string sysAdmin = "System Administrator";

            if (!userRoles.Any(entity => entity["name"].ToString() == sysAdmin))
            {
                Entity role = GetRoleByName(sysAdmin, businessUnit.Id);
                EntityReferenceCollection erc = new EntityReferenceCollection() { role.ToEntityReference() };

                AddRolesToUser(user.Id, erc);
            }
        }

        public static SetBusinessSystemUserResponse SetUserBusinessUnit(Entity user, Entity businessUnit)
        {
            SetBusinessSystemUserRequest bsRequest = new SetBusinessSystemUserRequest
            {
                BusinessId = businessUnit.Id,
                UserId = user.Id,
                ReassignPrincipal = user.ToEntityReference()
            };

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                OrganizationResponse response = service.Execute(bsRequest);
                SetBusinessSystemUserResponse bsResponse = (SetBusinessSystemUserResponse)response;

                return bsResponse;
            }
        }

        public static void MakeUserCda(string email)
        {
            Entity user = GetUserByEmail(email);

            string desiredBusinessUnit = "Corporate";
            Entity businessUnit = GetBusinessUnitByName(desiredBusinessUnit);

            string desiredTeam = "CDA";
            Entity team = GetTeamByName(desiredTeam);

            //change business unit
            SetUserBusinessUnit(user, businessUnit);

            //remove current roles to prevent role already exists error
            DataCollection<Entity> userRoles = FindUserRoles(user.Id);
            EntityReferenceCollection erc = new EntityReferenceCollection();
            erc.AddRange(userRoles.Select(a => a.ToEntityReference()));
            RemoveRolesFromUser(erc, user);

            EntityReferenceCollection roles = new EntityReferenceCollection()
            {
                GetRoleByName("DAI - CDA Full Access", businessUnit).ToEntityReference(),
                GetRoleByName("DAI - Contract Received Admin", businessUnit).ToEntityReference()
            };

            //set roles
            AddRolesToUser(user, roles);

            //add to team
            AddUsersToTeam(new[] { user }, team);
        }

        public static EntityCollection TT()
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("account")
                {
                    ColumnSet = new ColumnSet(true)
                };
                LinkEntity link = query.AddLink("plus_account_plus_orgtypes", "accountid", "plus_orgtypesid");
                //link.LinkCriteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, userId));

                return service.RetrieveMultiple(query);
            }
        }

        public static Entity Get(string entityName, Guid id)
        {
            return Get(entityName, id, new ColumnSet(true));
        }

        public static Entity Get(string entityName, Guid id, ColumnSet columnSet)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                try
                {
                    return service.Retrieve(entityName, id, columnSet);
                }
                catch 
                {
                    return null;
                }
            }
        }

        public static void Update(Entity entity)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                service.Update(entity);
            }
        }

        public static Guid Create(Entity entity)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                return service.Create(entity);
            }
        }

        public static Entity GetOrganizationByName(string name)
        {
            ConditionExpression condition = new ConditionExpression("account", "name", ConditionOperator.Equal, name);
            QueryExpression query = new QueryExpression("account");
            query.Criteria.AddCondition(condition);

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }

        public static Entity GetSectorByName(string name)
        {
            ConditionExpression condition = new ConditionExpression("plus_daisector", "plus_name", ConditionOperator.Equal, name);
            QueryExpression query = new QueryExpression("plus_daisector");
            query.Criteria.AddCondition(condition);
            

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }

        public static Entity GetBuyInClientOperatingUnit(string name)
        {
            ConditionExpression condition = new ConditionExpression("dai_buyinclientoperatingunit", "dai_name", ConditionOperator.Equal, name);
            QueryExpression query = new QueryExpression("dai_buyinclientoperatingunit");
            query.Criteria.AddCondition(condition);


            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }

        public static Entity GetPracticeAreaByName(Guid sectorGuid, string practiceArea)
        {
            ConditionExpression condition = new ConditionExpression("dai_practicearea", "dai_name", ConditionOperator.Equal, practiceArea);
            ConditionExpression condition2 = new ConditionExpression("dai_practicearea", "dai_daisectorid", ConditionOperator.Equal, sectorGuid);
            QueryExpression query = new QueryExpression("dai_practicearea");
            query.Criteria.AddCondition(condition);
            query.Criteria.AddCondition(condition2);
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }
        public static DataCollection<Entity> GetCdaWorkflows()
        {
            ConditionExpression condition = new ConditionExpression("dai_cda", "bpf_name", ConditionOperator.Equal, "CDA - Business Process Flow");
            QueryExpression query = new QueryExpression("dai_cda")
            {
                ColumnSet = new ColumnSet(true)
            };
            query.Criteria.AddCondition(condition);
            //bpf_opportunityid
            //activestageid

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities;
            }
        }
        public static DataCollection<Entity> FindWorkflows(string workflowEntity)
        {
            QueryExpression query = new QueryExpression(workflowEntity)
            {
                ColumnSet = new ColumnSet(true)
            };

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities;
            }
        }

        public static Entity GetWorkflowInstance(string workflowEntity, Entity entity)
        {
            QueryExpression query = new QueryExpression(workflowEntity)
            {
                ColumnSet = new ColumnSet(true)
            };

            if (entity != null)
            {
                string fieldName = string.Empty;
                if (entity.LogicalName == "opportunity")
                {
                    fieldName = "bpf_opportunityid";
                }
                else if (entity.LogicalName == "plus_project")
                {
                    fieldName = "bpf_plus_projectid";
                }

                ConditionExpression condition = new ConditionExpression(workflowEntity, fieldName, ConditionOperator.Equal, entity.Id);
                query.Criteria.AddCondition(condition);
            }

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }
        public static void Delete(Entity entity)
        {
            Delete(entity.LogicalName, entity.Id);
        }
        public static void Delete(string entityName, Guid id)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                service.Delete(entityName, id);
            }
        }
        public static void DeleteInstanceIfExists(string entityWorkflow, Entity entity)
        {
            var bpf = GetWorkflowInstance(entityWorkflow, entity);
            if (bpf != null)
            {
                Delete(bpf);
            }
        }

        public static OrganizationResponse AssignOwner(Entity owner, Entity target)
        {
            
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                AssignRequest assign = new AssignRequest
                {
                    Assignee = owner.ToEntityReference(),
                    Target = target.ToEntityReference()
                };

                return service.Execute(assign);
            }
        }
        public static DataCollection<Entity> FindOpportunityTechTags(Guid id)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression oppTechTags = new QueryExpression("plus_technicalscope")
                {
                    ColumnSet = new ColumnSet(new string[] { "plus_projectid", "plus_proposalbidid" })
                };
                oppTechTags.Criteria.AddCondition(new ConditionExpression("plus_proposalbidid", ConditionOperator.Equal, id));
                oppTechTags.Criteria.AddCondition(new ConditionExpression("plus_projectid", ConditionOperator.Null));

                return service.RetrieveMultiple(oppTechTags).Entities;
            }
        }

        public static DataCollection<Entity> FindAkaForOpportunity(Guid id)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("plus_aka")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("plus_aka", "dai_proposalbidid", ConditionOperator.Equal, id));
                EntityCollection result = service.RetrieveMultiple(query);

                return result.Entities;
            }
        }
        public static Entity GetClientsByName(string name)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("account")
                {
                    ColumnSet = new ColumnSet(true)
                };
            query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, new[] { name }));
                EntityCollection result = service.RetrieveMultiple(query);

                return result.Entities[0];
            }
            
        }
        public static Entity GetCurrencyByIsoCode(string currencyCode)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("transactioncurrency")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("isocurrencycode", ConditionOperator.Equal, new[] { currencyCode}));
                EntityCollection result = service.RetrieveMultiple(query);

                return result.Entities[0];
            }
        }

        public static Entity GetProjectCurrencyByIsoCode(string currencyCode)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("isi_projectcurrency")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("isi_currencycode", ConditionOperator.Equal, new[] { currencyCode }));
                EntityCollection result = service.RetrieveMultiple(query);

                return result.Entities[0];
            }
        }

        public static Entity GetContactByEmail(string email)
        {
            DataCollection<Entity> contacts = FindContactByEmail(new[] { email });
            return contacts.Any() ? contacts.First() : null;
        }

        public static DataCollection<Entity> FindContactByEmail(string email)
        {
            return FindContactByEmail(new[] {email });
        }

        public static DataCollection<Entity> FindContactByEmail(string[] emails)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                FilterExpression filter = new FilterExpression(LogicalOperator.Or);

                foreach (var email in emails)
                {
                    filter.AddCondition("emailaddress1", ConditionOperator.Equal, new[] { email });
                    filter.AddCondition("emailaddress2", ConditionOperator.Equal, new[] { email });
                    filter.AddCondition("emailaddress3", ConditionOperator.Equal, new[] { email }); 
                }

                QueryExpression query = new QueryExpression("contact")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = filter
                };

                EntityCollection result = service.RetrieveMultiple(query);

                return result.Entities;
            }
        }

        public static void AddThemesToProjectSettings(Guid projectSettingsGuid, EntityReferenceCollection themes)
        {
            if (themes == null || !themes.Any()) return;

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                try
                {
                    service.Associate("new_grantprogram", projectSettingsGuid, new Relationship("isi_CrosscuttingTheme_Project_New_grantpr"), themes);
                }
                catch (Exception e)
                {
                    Log<DynamicsHelper>.Logger.Error(e);
                }
            }
        }

        public static Entity FindCountryByName(string country)
        {
            if (string.IsNullOrWhiteSpace(country)) return default;

            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("plus_country")
                {
                    ColumnSet = new ColumnSet(true),
                    TopCount = 1
                };

                query.Criteria.AddCondition(new ConditionExpression("plus_name", ConditionOperator.Equal, new[] { country }));

                EntityCollection result = service.RetrieveMultiple(query);

                return result.Entities.Any() ? result.Entities[0] : null;
            }
        }

        public static void AddCountriesToConcept(Guid conceptGuid, EntityReferenceCollection countries)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                try
                {
                    service.Associate("new_proposal", conceptGuid, new Relationship("dai_New_proposal_plus_country_plus_country"), countries);
                }
                catch (Exception e)
                {
                    Log<DynamicsHelper>.Logger.Error(e);
                }
            }
        }
        public static void AddCountriesToProject(Guid ProjectId, EntityReferenceCollection countries)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                try
                {
                    service.Associate("plus_project",ProjectId, new Relationship("plus_plus_project_plus_country"), countries);
                }
                catch (Exception e)
                {
                    Log<DynamicsHelper>.Logger.Error(e);
                }
            }
        }

        public static void AddClientToProject(Guid projectGuid, Guid clientGuid)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                try
                {
                    service.Associate("plus_project", projectGuid, new Relationship("plus_plus_project_plus_clientcontact"), new EntityReferenceCollection() { new EntityReference("plus_clientcontact", clientGuid) });
                }
                catch (Exception e)
                {
                    Log<DynamicsHelper>.Logger.Error(e);
                }
            }
        }

        public static void AddUserToDAITeamForProject(Guid projectGuid, Guid DAITeamMember)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                try
                {
                    service.Associate("plus_project", projectGuid, new Relationship("plus_plus_project_plus_daiteam"), new EntityReferenceCollection() { new EntityReference("plus_daiteam", DAITeamMember) });
                }
                catch (Exception e)
                {
                    Log<DynamicsHelper>.Logger.Error(e);
                }
            }
        }
        public static DataCollection<Entity> FindOrganizationByName(string name)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                FilterExpression filter = new FilterExpression(LogicalOperator.Or);
                filter.AddCondition("name", ConditionOperator.Equal, new[] { name });
                filter.AddCondition("plus_legalname", ConditionOperator.Equal, new[] { name });

                QueryExpression query = new QueryExpression("account")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = filter
                };

                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities;
            }
        }

        public static DataCollection<Entity> FindDocumentTypesForProjectSetting(Guid projectSettingId)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("isi_documenttypesforproject")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("isi_projectsettingsid", ConditionOperator.Equal, projectSettingId ));
                query.Criteria.AddCondition(new ConditionExpression("isi_stagee", ConditionOperator.Equal, (int)DocumentStage.Concept));

                EntityCollection result = service.RetrieveMultiple(query);
                return result.Entities;
            }
        }

        public static Entity GetProjectDocTypeFromProjectSetting(Guid projectSetting, Guid documentType)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("isi_documenttypesforproject")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("isi_projectsettingsid", ConditionOperator.Equal, projectSetting));
                query.Criteria.AddCondition(new ConditionExpression("isi_documenttypeid", ConditionOperator.Equal, documentType));
                query.Criteria.AddCondition(new ConditionExpression("isi_stagee", ConditionOperator.Equal, (int)DocumentStage.Concept));

                EntityCollection result = service.RetrieveMultiple(query);
                if (result != null && result.Entities.Any()) return result.Entities[0];
                return default;
            }
        }

        public static Entity FindConceptByTitle(string title)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("new_proposal")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("new_name", ConditionOperator.Equal, title));

                EntityCollection result = service.RetrieveMultiple(query);
                if (result != null && result.Entities.Any()) return result.Entities[0];
                return default;
            }
        }

        public static Entity FindSolicitationByNumber(string solicitationNumber)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("new_foa")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("dai_numberofsolicitation", ConditionOperator.Equal, solicitationNumber));

                EntityCollection result = service.RetrieveMultiple(query);
                if (result != null && result.Entities.Any()) return result.Entities[0];
                return default;
            }
        }

        public static Entity FindTeamByName(string name)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("team")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, name));

                EntityCollection result = service.RetrieveMultiple(query);
                if (result != null && result.Entities.Any()) return result.Entities[0];
                return default;
            }
        }

        public static Entity GetProjectByNameAndOracleId(string name, string oracleId)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken= AccessToken2() })
            {
                QueryExpression query = new QueryExpression("plus_project")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("plus_name", ConditionOperator.Equal, name));
                query.Criteria.AddCondition(new ConditionExpression("plus_oracleprojectnumber", ConditionOperator.Equal, oracleId));

                EntityCollection result = service.RetrieveMultiple(query);
                if (result != null && result.Entities.Any()) return result.Entities[0];
                return default;
            }
        }

        public static Entity FindProjectSettingByName(string name)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("new_grantprogram")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("new_name", ConditionOperator.Equal, name));

                EntityCollection result = service.RetrieveMultiple(query);
                if (result != null && result.Entities.Any()) return result.Entities[0];
                return default;
            }
        }

        public static Entity FindConceptByFormsiteRefNum(string refNum)
        {
            using (OrganizationWebProxyClient service = new OrganizationWebProxyClient(_orgService, new TimeSpan(0, 5, 0), false) { HeaderToken = AccessToken2() })
            {
                QueryExpression query = new QueryExpression("new_proposal")
                {
                    ColumnSet = new ColumnSet(true)
                };
                query.Criteria.AddCondition(new ConditionExpression("dai_formsitereferencenumber", ConditionOperator.Equal, refNum));

                EntityCollection result = service.RetrieveMultiple(query);
                if (result != null && result.Entities.Any()) return result.Entities[0];
                return default;
            }
        }
    }
}
