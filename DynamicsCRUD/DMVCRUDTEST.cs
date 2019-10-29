using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Client;

using System.Web.Services.Description;
using Microsoft.VisualBasic.FileIO;
using System.Web;


using System.ServiceModel;
using System.ServiceModel.Description;



// These namespaces are found in the Microsoft.Xrm.Sdk.dll assembly
// found in the SDK\bin folder.

using Microsoft.Xrm.Sdk.Discovery;



namespace DynamicsCRUD
{
    class DMVCRUDTEST
    {
        static string strConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["MyCDSServer"].ToString();
        static CrmServiceClient svc = new CrmServiceClient(strConnectionString);

        static void Main(string[] args)
        {
            Console.WriteLine("Dynamics CRM connected");

            int userInput;
            do
            {
                userInput = DisplayMainMenu();

                if (userInput == 1)
                    ReadDMVAppsFromCSV();
                else if (userInput == 2)
                    DeleteAppsCreatedToday();
                else if (userInput == 3)
                    RetrieveAppsCreatedToday();
                else if (userInput == 4)
                    PCRetrieveCustWithAppInTx();
                else if (userInput == 5)
                    CreateCustomer();
                else if (userInput == 6)
                    UpdateAppsInclSubject();
                else if (userInput == 7)
                    DeleteSpecificApplication();
                else if (userInput == 8)
                    ListVehRegApps();
                else if (userInput == 9)
                    CreateAnAppWithValues();
                else if (userInput == 10)
                    AssignAccessTeam();

                Console.WriteLine("Process Finished");
            } while (userInput != 0);
        }

        static public int DisplayMainMenu()
        {
            Console.WriteLine();
            Console.WriteLine("CRUD Menu:");
            Console.WriteLine("1. Read from CSV and create Applications");
            Console.WriteLine("2. Delete Applications created today");
            Console.WriteLine("3. List Applications created today");
            Console.WriteLine("4. List all Customers in TX with applications");
            Console.WriteLine("5. Create a Customer with name and SSN");
            Console.WriteLine("6. Update all Applications without a subject to have a generated subject");
            Console.WriteLine("7. Delete a specific Application from today");
            Console.WriteLine("8. List all Apps of type Vehicle Registration");
            Console.WriteLine("9. Create an Application by selecting type and name");
            Console.WriteLine("10. Assign application to team");
            Console.WriteLine("0. Exit");
            var result = Console.ReadLine();
            return Convert.ToInt32(result);
        }




        public static void AssignAccessTeam()
        {


            Entity entTeam = new Entity("team");
            entTeam.Attributes["name"] = "test unit";
            entTeam["businessunitid"] = new EntityReference("businessunit", new Guid("7DFC241C-06DA-E911-AA24-000D3A3ABDB6"));
            Guid teamId = svc.Create(entTeam);


            Console.WriteLine("Created {0}", entTeam.Attributes["name"]);

            // Add the role to the team.
            svc.Associate(
                       "team",
                       teamId,
                       new Relationship("teamroles_association"),
                       new EntityReferenceCollection() { new EntityReference("role", new Guid("81FC241C-06DA-E911-AA24-000D3A3ABDB6")) });

            Console.WriteLine("Assigned team to role");

            // It takes some time for the privileges to propagate to the team. Delay the
            // application until the privilege has been assigned.
            bool teamLacksPrivilege = true;
            while (teamLacksPrivilege)
            {
                RetrieveTeamPrivilegesRequest retrieveTeamPrivilegesRequest =
                    new RetrieveTeamPrivilegesRequest
                    {
                        TeamId = teamId
                    };

                RetrieveTeamPrivilegesResponse retrieveTeamPrivilegesResponse =
                    (RetrieveTeamPrivilegesResponse)svc.Execute(
                    retrieveTeamPrivilegesRequest);

                foreach (RolePrivilege rp in
                    retrieveTeamPrivilegesResponse.RolePrivileges)
                {
                    if (rp.PrivilegeId == new Guid("81FC241C-06DA-E911-AA24-000D3A3ABDB6"))
                    {
                        teamLacksPrivilege = false;
                        break;
                    }
                    else
                    {
                        System.Threading.Thread.CurrentThread.Join(500);
                    }
                }
            }

                AssignRequest assignRequest = new AssignRequest()
            {
                Assignee = new EntityReference("team", teamId),
                Target = new EntityReference("madmv_ma_application", new Guid("6d52f446-38ec-e911-a812-000d3a3349d4"))  // app 1469
            };
        
 

            svc.Execute(assignRequest);

            Console.WriteLine("The account is owned by the team.");


        }

        public static void UpdateAppsInclSubject()
        {
            string FetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='madmv_ma_application'>
                                    <attribute name='madmv_ma_applicationid' />
                                    <attribute name='madmv_appid' />
                                    <attribute name='madmv_applicationtype' />
                                    <attribute name='madmv_ownerinfo' />
                                    <order attribute='madmv_appid' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='madmv_applicationsubject' operator='null' />
                                    </filter>
                                  </entity>
                                </fetch>";

            FetchXML = string.Format(FetchXML);
            EntityCollection data = svc.RetrieveMultiple(new FetchExpression(FetchXML));

            foreach (Entity ent in data.Entities)
            {
                string myLookupName = ((EntityReference)ent.Attributes["madmv_ownerinfo"]).Name;
                string appType = ent.FormattedValues["madmv_applicationtype"].ToString();
                Guid Id = ent.Id;
                ent.Id = Id;
                ent.Attributes["madmv_applicationsubject"] = myLookupName + " " + appType;
                svc.Update(ent);
            }
        }


        public static void DeleteSpecificApplication()
        {
            EntityCollection data = RetrieveAppsCreatedToday();  //gets and list all apps created today with a number before to choose
            int numApps = 1; // counter int to check if selected app is the one in the loop
            var validChoice = false;

            Console.WriteLine("Select the number corresponding to the Application you wish to delete or 0 to quit.");
            var chosenNum = Console.ReadLine();
            if (chosenNum == "0")
                return;


            foreach (Entity ent in data.Entities)
            {
                Guid Id = ent.Id;
                if (int.Parse(chosenNum) == numApps)
                {
                    while (!validChoice)
                    {
                        validChoice = true;
                        Console.WriteLine("Are you sure you wish to delete Application " + ent.Attributes["madmv_appid"] + "? (y/n)");
                        var choice = Console.ReadLine();
                        if (choice.ToLower() == "n")
                            return;
                        else if (choice.ToLower() == "y")
                        {
                            Console.WriteLine("Processing Delete Action\n...");
                            svc.Delete("madmv_ma_application", Id);
                            Console.WriteLine("Application Deleted!");
                        }
                        else
                        {
                            Console.WriteLine("Invalid choice!");
                            validChoice = false;
                        }
                    }
                }
                numApps++;
            }
        }

        public static void DeleteAppsCreatedToday()
        {
            EntityCollection data = RetrieveAppsCreatedToday();
            var validChoice = false;
            int index = 0;

            while (!validChoice)
            {
                validChoice = true;
                Console.WriteLine("Are you sure you wish to delete the above applications? (y/n)");
                var choice = Console.ReadLine();
                if (choice.ToLower() == "n")
                    return;
                else if (choice.ToLower() == "y")
                {
                    Console.WriteLine("Processing Delete Action\n...");
                    foreach (Entity ent in data.Entities)
                    {
                        Guid Id = ent.Id;
                        svc.Delete("madmv_ma_application", Id);
                        index++;
                    }
                    Console.WriteLine(index + " Application(s) Deleted!");
                } else
                {
                    Console.WriteLine("Invalid choice!");
                    validChoice = false;
                }
            }

        }

        public static void CreateCustomer()
        {
            Console.WriteLine("Please enter the name for the Customer or 0 to exit (firstName lastName)");
            var name = Console.ReadLine();
            if (name == "0")
                return;

            Console.WriteLine("Please enter the 9 digit SSN or 0 to cancel out");
            var ssn = Console.ReadLine();
            if (name == "0")
                return;

            Entity entCustomer = new Entity("madmv_ma_customer");

            entCustomer.Attributes["madmv_fullname"] = name;
            entCustomer.Attributes["madmv_cssn"] = ssn;
            string[] nameSplit = name.Split(' ');  // TODO check if full name has more then 2 names and deal with it.
            entCustomer.Attributes["madmv_firstname"] = nameSplit[0];
            entCustomer.Attributes["madmv_lastname"] = nameSplit[1];  // TODO check if name had a space to avoid out of bounds issue
            svc.Create(entCustomer);
        }

        public static void CreateAnAppWithValues()
        {
            Entity entApplication = new Entity("madmv_ma_application");
            int numCust = 1;

            Console.WriteLine("Select the number corresponding to the desired Application type or 0 to exit (ex. 1)");
            Console.WriteLine("1. Vehicle Registration");
            Console.WriteLine("2. Address Change");
            var typeChoice = Console.ReadLine();
            if (typeChoice == "0")
                return;
            else if (typeChoice == "1")
                entApplication["madmv_applicationtype"] = new OptionSetValue(876570000);
            else if (typeChoice == "2")
                entApplication["madmv_applicationtype"] = new OptionSetValue(876570001);

            EntityCollection data = RetrieveCustomerList();
            Console.WriteLine("Select the number corresponding to the Customer you wish to assign to the Application or 0 to quit.");
            var chosenNum = Console.ReadLine();
            if (chosenNum == "0")
                return;

            foreach (Entity ent in data.Entities)
            {
                Guid Id = ent.Id;
                if (int.Parse(chosenNum) == numCust)
                {
                    entApplication["madmv_ownerinfo"] = new EntityReference("madmv_ma_customer", Id);
                }
                numCust++;
            }

            svc.Create(entApplication);
        }

        public static void CreateNewContact(String name, String ssn)
        {
            Entity newentApplication = new Entity("madmv_ma_customer");

            newentApplication.Attributes["madmv_fullname"] = name;
            newentApplication.Attributes["madmv_cssn"] = ssn;
            string[] nameSplit = name.Split(' ');  // TODO check if full name has more then 2 names and deal with it.
            newentApplication.Attributes["madmv_firstname"] = nameSplit[0];
            newentApplication.Attributes["madmv_lastname"] = nameSplit[1];  // TODO check if name had a space to avoid out of bounds issue
            svc.Create(newentApplication);
        }

        public static void CreateNewVehicle(String vin)
        {
            Entity newentApplication = new Entity("madmv_ma_vehicle");

            newentApplication.Attributes["madmv_vehicleidnumber"] = vin;
            svc.Create(newentApplication);
        }

        public static void CreateNewCompany(String name)
        {
            Entity newentApplication = new Entity("madmv_ma_company");

            newentApplication.Attributes["madmv_name"] = name;
            svc.Create(newentApplication);
        }

        public static void ReadDMVAppsFromCSV()
        {
            var index = 0;

            using (TextFieldParser parser = new TextFieldParser(@"C:\Users\matth\Documents\dmvappcsv.csv"))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.ReadLine();

                while (!parser.EndOfData)
                {
                    Entity entApplication = new Entity("madmv_ma_application");
                    //Processing row
                    string[] fields = parser.ReadFields();

                    var appType = fields[0];
                    var street1 = fields[1];
                    var street2 = fields[2];
                    var city = fields[3];
                    var state = fields[4];
                    var zip = fields[5];
                    var country = fields[6];
                    var reasonChange = fields[7];
                    var ssn = fields[19];
                    

                    DateTime cEffectiveDate = DateTime.Now;

                    var effectiveDate = fields[8];
                    if (effectiveDate != "")
                        cEffectiveDate = Convert.ToDateTime(effectiveDate);



                    var ownerInfo = fields[9];
                    EntityCollection ownerData = RetrieveCustomerMatchingSSN(ssn);
                    if (ownerData.Entities.Count > 0)
                    {
                        foreach (Entity ent in ownerData.Entities)
                        {
                            Guid Id = ent.Id;
                            String ssnToCompare = ent.Attributes["madmv_cssn"].ToString();
                            entApplication["madmv_ownerinfo"] = new EntityReference("madmv_ma_customer", Id);
                        }
                    }
                    else if (ownerInfo != "")
                    {
                        Console.WriteLine("Creating Customer " + ownerInfo);
                        CreateNewContact(ownerInfo, ssn);

                        EntityCollection newOwnerData = RetrieveCustomerMatchingSSN(ssn);
                        foreach (Entity ent in newOwnerData.Entities)
                        {
                            Guid Id = ent.Id;
                            entApplication["madmv_ownerinfo"] = new EntityReference("madmv_ma_customer", Id);
                        }
                    }


                    var vehicleDetails = fields[10];
                    EntityCollection vehicleData = RetrieveVehiclesMatchingVin(vehicleDetails);
                    if (vehicleData.Entities.Count > 0)
                    {
                        foreach (Entity ent in vehicleData.Entities)
                        {
                            Guid Id = ent.Id;
                            String vinToCompare = ent.Attributes["madmv_vehicleidnumber"].ToString();
                            entApplication["madmv_vehicledetails"] = new EntityReference("madmv_ma_vehicle", Id);
                        }
                    }
                    else if (vehicleDetails != "")
                    {
                        Console.WriteLine("Creating Vehicle " + vehicleDetails);
                        CreateNewVehicle(vehicleDetails);

                        EntityCollection newVehicleData = RetrieveVehiclesMatchingVin(vehicleDetails);
                        foreach (Entity ent in newVehicleData.Entities)
                        {
                            Guid Id = ent.Id;
                            entApplication["madmv_vehicledetails"] = new EntityReference("madmv_ma_vehicle", Id);
                        }
                    }


                    var lesseeCompany = fields[11];
                    EntityCollection companyData = RetrieveCompanyMatchingName(lesseeCompany);
                    if (companyData.Entities.Count > 0)
                    {
                        foreach (Entity ent in companyData.Entities)
                        {
                            Guid Id = ent.Id;
                            String nameToCompare = ent.Attributes["madmv_name"].ToString();
                            entApplication["madmv_lesseecompany"] = new EntityReference("madmv_ma_company", Id);
                        }
                    }
                    else if (lesseeCompany != "")
                    {
                        Console.WriteLine("Creating Company " + lesseeCompany);
                        CreateNewCompany(lesseeCompany);

                        EntityCollection newCompanyData = RetrieveCompanyMatchingName(lesseeCompany);
                        foreach (Entity ent in newCompanyData.Entities)
                        {
                            Guid Id = ent.Id;
                            entApplication["madmv_lesseecompany"] = new EntityReference("madmv_ma_company", Id);
                        }
                    }

                    var registrationPeriod = fields[12];
                    var registrationType = fields[13];
                    var describeOther = fields[14];
                    var plateType = fields[15];
                    var reissuedPlate = fields[16];
                    var insuranceCompany = fields[17];
                    var fee = fields[18];


                    if (appType != "")
                        entApplication["madmv_applicationtype"] = new OptionSetValue(int.Parse(appType));
                    if (registrationPeriod != "")
                        entApplication["madmv_registrationperiod"] = new OptionSetValue(int.Parse(registrationPeriod));
                    if (registrationType != "")
                        entApplication["madmv_registrationtype"] = new OptionSetValue(int.Parse(registrationType));
                    if (plateType != "")
                        entApplication["madmv_platetype"] = new OptionSetValue(int.Parse(plateType));
                    if (reissuedPlate != "")
                        entApplication["madmv_reissuedplates"] = new OptionSetValue(int.Parse(reissuedPlate));
                    if (fee != "")
                        entApplication.Attributes["madmv_fee"] = int.Parse(fee);

                    entApplication.Attributes["madmv_newstreet1"] = street1;
                    entApplication.Attributes["madmv_newstreet2"] = street2;
                    entApplication.Attributes["madmv_newcity"] = city;
                    entApplication.Attributes["madmv_newstate"] = state;
                    entApplication.Attributes["madmv_newzip"] = zip;
                    entApplication.Attributes["madmv_newcountry"] = country;
                    entApplication.Attributes["madmv_reasonforaddresschange"] = reasonChange;
                    entApplication.Attributes["madmv_addressfieldeffectivedate"] = cEffectiveDate;
                    entApplication.Attributes["madmv_describeother"] = describeOther;
                    entApplication.Attributes["madmv_insurancecompany"] = insuranceCompany;
                    entApplication.Attributes["madmv_ssn"] = ssn;

                    index++;
                    if (int.Parse(appType) == 876570000)
                        Console.WriteLine("Creating Application " + index + ": App Type- Vehicle Registration    Customer- " + ownerInfo);
                    else if (int.Parse(appType) == 876570001)
                        Console.WriteLine("Creating Application " + index + ": App Type- Address Change    Customer- " + ownerInfo);
                    svc.Create(entApplication);
                }
            }
        }

        public static void ListVehRegApps()
        {
            //  Query using ConditionExpression and FilterExpression
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "madmv_applicationtype";
            condition1.Operator = ConditionOperator.Like;
            condition1.Values.Add(876570000);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);
            QueryExpression query = new QueryExpression("madmv_ma_application");

            query.ColumnSet.AddColumns("madmv_appid");
            query.ColumnSet.AddColumns("createdon");
            query.ColumnSet.AddColumns("statuscode");
            query.ColumnSet.AddColumns("madmv_applicationtype");

            query.Criteria.AddFilter(filter1);

            EntityCollection result1 = svc.RetrieveMultiple(query);
            Console.WriteLine(); Console.WriteLine("Query using Query Expression with ConditionExpression and FilterExpression");
            Console.WriteLine("---------------------------------------");
            foreach (var a in result1.Entities)
            {
                Console.WriteLine("Application ID: " + a.Attributes["madmv_appid"] + "      " + a.Attributes["createdon"] + "      " + a.FormattedValues["statuscode"] + "      " + a.FormattedValues["madmv_applicationtype"]);
            }
            Console.WriteLine("---------------------------------------");
        }

        public static EntityCollection RetrieveCustomerList()
        {
            //  Query using ConditionExpression and FilterExpression
            QueryExpression query = new QueryExpression("madmv_ma_customer");
            query.ColumnSet.AddColumns("madmv_fullname");
            int count = 1;

            EntityCollection data = svc.RetrieveMultiple(query);

            Console.WriteLine("---------------------------------------");
            for (int intCount = 0; intCount < data.Entities.Count; intCount++)
            {
                Console.WriteLine(count + ")  Customer Name : " + data.Entities[intCount].Attributes["madmv_fullname"]);
                count++;
            }
            return data;
        }

        public static EntityCollection RetrieveCustomerMatchingSSN(string ssn)
        {
            //  Query using ConditionExpression and FilterExpression
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "madmv_cssn";
            condition1.Operator = ConditionOperator.Like;
            condition1.Values.Add(ssn);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);

            QueryExpression query = new QueryExpression("madmv_ma_customer");
            query.ColumnSet.AddColumns("madmv_cssn");
            query.Criteria.AddFilter(filter1);

            return svc.RetrieveMultiple(query);
        }

        public static EntityCollection RetrieveVehiclesMatchingVin(string vin)
        {
            //  Query using ConditionExpression and FilterExpression
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "madmv_vehicleidnumber";
            condition1.Operator = ConditionOperator.Like;
            condition1.Values.Add(vin);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);

            QueryExpression query = new QueryExpression("madmv_ma_vehicle");
            query.ColumnSet.AddColumns("madmv_vehicleidnumber");
            query.Criteria.AddFilter(filter1);

            return svc.RetrieveMultiple(query);
        }

        public static EntityCollection RetrieveCompanyMatchingName(string name)
        {
            //  QueryExpression to find matching company records and return as EntityCollection
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "madmv_name";
            condition1.Operator = ConditionOperator.Like;
            condition1.Values.Add(name);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);

            QueryExpression query = new QueryExpression("madmv_ma_company");
            query.ColumnSet.AddColumns("madmv_name");
            query.Criteria.AddFilter(filter1);

            return svc.RetrieveMultiple(query);
        }
    
        public static EntityCollection RetrieveAppsCreatedToday()
        {
            Entity entApplication = new Entity("madmv_ma_application");
            int count = 1;

            string FetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='madmv_ma_application'>
                                    <attribute name='madmv_ownerinfo' />
                                    <attribute name='madmv_applicationtype' />
                                    <attribute name='madmv_ma_applicationid' />
                                    <attribute name='madmv_appid' />
                                    <attribute name='createdon' />
                                    <order attribute='madmv_appid' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='createdon' operator='today' />
                                    </filter>
                                  </entity>
                                </fetch>";

            FetchXML = string.Format(FetchXML);
            EntityCollection data = svc.RetrieveMultiple(new FetchExpression(FetchXML));


            Console.WriteLine("---------------------------------------");
            for (int intCount = 0; intCount < data.Entities.Count; intCount++)
            {
                string myLookupName = "N/A";
                if (((EntityReference)data.Entities[intCount].Attributes["madmv_ownerinfo"]) != null)
                {
                     myLookupName = ((EntityReference)data.Entities[intCount].Attributes["madmv_ownerinfo"]).Name;
                }
                Console.WriteLine(count + ")  Application ID : " + data.Entities[intCount].Attributes["madmv_appid"].ToString() + "    Application Type : " + data.Entities[intCount].FormattedValues["madmv_applicationtype"].ToString() + "    Owner : " + myLookupName);
                count++;
            }
            Console.WriteLine("---------------------------------------");
            return data;
        }

        public static void PCRetrieveCustWithAppInTx()
        {
            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = "madmv_ma_application",
                ColumnSet = new ColumnSet("madmv_appid", "madmv_applicationtype"),
                LinkEntities =
                {
                    new LinkEntity
                    {
                        EntityAlias = "custEnt",
                        JoinOperator = JoinOperator.Inner,
                        LinkFromAttributeName = "madmv_ssn",
                        LinkFromEntityName = "madmv_ma_application",
                        LinkToAttributeName = "madmv_cssn",
                        LinkToEntityName = "madmv_ma_customer",
                        Columns = new ColumnSet("madmv_fullname", "madmv_stateprovince"),//select fields for Entity B
                        LinkCriteria =
                        {
                            Conditions =
                            {
                                new ConditionExpression("madmv_stateprovince", ConditionOperator.Equal, "TX")
                            }
                        }
                    }
                },
                Criteria =
                {
                    Filters =
                    {
                        new FilterExpression
                        {
                            FilterOperator = LogicalOperator.And,
                            Conditions =
                            {
                                new ConditionExpression("madmv_applicationtype", ConditionOperator.Equal, "876570001"),
                            },
                        }
                    }
                }
            };

            // DataCollection<Entity> entityCollection = svc.RetrieveMultiple(query).Entities;
            EntityCollection result1 = svc.RetrieveMultiple(query);
            // Display the results.
            Console.WriteLine("List all Customers matching specified parameters");
            Console.WriteLine("---------------------------------------");
            foreach (Entity a in result1.Entities)
            {
                Console.WriteLine("Customer Name: {0}", ((AliasedValue)a.Attributes["custEnt.madmv_fullname"]).Value);
                Console.WriteLine("State: {0}", ((AliasedValue)a.Attributes["custEnt.madmv_stateprovince"]).Value);
                Console.WriteLine("Application ID {0}", a.Attributes["madmv_appid"]);
                Console.WriteLine("Application Type: {0}", a.FormattedValues["madmv_applicationtype"].ToString());
            }
            Console.WriteLine("---------------------------------------");
            Console.WriteLine();
        }
    }
}



