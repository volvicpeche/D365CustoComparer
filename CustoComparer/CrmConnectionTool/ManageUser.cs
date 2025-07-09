using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
namespace CrmConnectionTool
{
    public class ManageUser
    {
        public static List<string> GetOrganizations(string DiscoverServiceURL, string UserName, string Password, string Domain)
        {
            ClientCredentials credentials = new ClientCredentials();

            credentials.Windows.ClientCredential = new System.Net.NetworkCredential(UserName, Password, Domain);
            using (var discoveryProxy = new DiscoveryServiceProxy(new Uri(DiscoverServiceURL), null, credentials, null))
            {
                discoveryProxy.Authenticate();

                // Get all Organizations using Discovery Service

                RetrieveOrganizationsRequest retrieveOrganizationsRequest =
                new RetrieveOrganizationsRequest()
                {
                    AccessType = EndpointAccessType.Default,
                    Release = OrganizationRelease.Current
                };

                RetrieveOrganizationsResponse retrieveOrganizationsResponse =
                (RetrieveOrganizationsResponse)discoveryProxy.Execute(retrieveOrganizationsRequest);

                if (retrieveOrganizationsResponse.Details.Count > 0)
                {
                    var orgs = new List<String>();

                    foreach (OrganizationDetail orgInfo in retrieveOrganizationsResponse.Details)
                        orgs.Add(orgInfo.UniqueName);

                    return orgs;
                }
                else
                    return null;
            }
        }

        public static void CreateUser(string username,string org, Uri oUri, string securityRoleName, ClientCredentials creds)
        {
            using (var service = new OrganizationServiceProxy(oUri, null, creds, null))
            {
                try
                {
                    //Get root business unit
                    var buRoot = service.RetrieveMultiple(new QueryExpression("businessunit")
                    {
                        NoLock = true,
                        ColumnSet = new ColumnSet("name"),
                        Criteria =
                            {
                                Conditions =
                                {
                                    new ConditionExpression("parentbusinessunitid", ConditionOperator.Null)
                                }
                            }
                    }).Entities.FirstOrDefault();

                    if (buRoot == null)
                    {
                        Console.Error.WriteLine($"Root Business unit not found in {oUri.ToString()}");
                        return;
                    }

                    //Check if user already exists
                    var monitoringUser = service.RetrieveMultiple(new QueryExpression("systemuser")
                    {
                        NoLock = true,
                        ColumnSet = new ColumnSet(),
                        Criteria =
                            {
                                Conditions =
                                {
                                    new ConditionExpression("domainname", ConditionOperator.Equal, username)
                                }
                            }
                    }).Entities.FirstOrDefault();

                    //if not create him
                    if (monitoringUser == null) 
                    {
                        monitoringUser = new Entity("systemuser")
                        {
                            ["domainname"] = username,
                            ["firstname"] = "Monitoring",
                            ["lastname"] = "service account",
                            ["businessunitid"] = buRoot.ToEntityReference()
                        };

                        monitoringUser.Id = service.Create(monitoringUser);
                    }
                    else
                    {
                        Console.WriteLine($"User already exists in {org}");
                    }

                    var securityRole = service.RetrieveMultiple(new QueryExpression("role")
                    {
                        ColumnSet = new ColumnSet(),
                        NoLock = true,
                        Criteria =
                        {
                            Conditions =
                                {
                                    new ConditionExpression("name",ConditionOperator.Equal, securityRoleName),
                                    new ConditionExpression("businessunitid", ConditionOperator.Equal, buRoot.Id)
                                }
                        }
                    }).Entities.FirstOrDefault();

                    if (securityRole != null)
                    {
                        try
                        {
                            service.Associate(
                                "systemuser",
                                monitoringUser.Id,
                                new Relationship("systemuserroles_association"),
                                new EntityReferenceCollection { securityRole.ToEntityReference() }
                            );
                        }
                        catch (Exception ex) {
                            Console.Error.WriteLine($"error associating security role {org} : {ex.Message}");
                            return;
                        }

                        

                        Console.WriteLine($"Sucess {oUri.ToString()}");
                    }
                    else
                    {
                        Console.Error.WriteLine($"{securityRoleName} not found in {oUri.ToString()}");
                        return;
                    }


                }
                catch (Exception ex) {
                    Console.Error.WriteLine($"Error : {ex.ToString()} in {oUri.ToString()}");
                    return;
                }
            }
        }
    }
}
