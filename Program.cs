using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using ClosedXML.Excel;

namespace CrmOpportunityProductReporter
{
    /// <summary>
    /// Main program class for generating opportunity product reports in Dynamics 365 CRM
    /// </summary>
    public class Program
    {
        // Cache for lookup values to minimize CRM queries
        private static readonly Dictionary<Guid, string> _lookupCache = new Dictionary<Guid, string>();

        // Cache for geography data to minimize CRM queries
        private static readonly Dictionary<Guid, List<string>> _geographyCache = new Dictionary<Guid, List<string>>();

        // Track execution start time for performance monitoring
        private static DateTime _startTime;

        /// <summary>
        /// Main entry point for the application
        /// </summary>
        public static void Main(string[] args)
        {
            try
            {
                _startTime = DateTime.Now;
                Console.WriteLine($"Report Generation Started at: {_startTime:yyyy-MM-dd HH:mm:ss}");

                RunReportGeneration();

                var endTime = DateTime.Now;
                Console.WriteLine($"\nReport Generation Completed at: {endTime:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"Total Execution Time: {(endTime - _startTime).TotalSeconds} seconds");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                Environment.Exit(1); // Exit with error code
            }
        }

        /// <summary>
        /// Main workflow for report generation
        /// </summary>
        private static void RunReportGeneration()
        {
            // Use using statement to ensure proper disposal of CRM service client
            using (var service = GetCrmServiceClient())
            {
                if (!service.IsReady)
                {
                    throw new Exception($"CRM connection failed: {service.LastCrmError}");
                }

                Console.WriteLine("Connected to CRM successfully");

                // Fetch users with specific filters
                var usersList = FetchFilteredUsers(service);

                // Process user hierarchy starting with HPR users
                ProcessUserHierarchy(service, usersList);
            }
        }

        /// <summary>
        /// Creates and configures the CRM service client using secure methods
        /// </summary>
        private static CrmServiceClient GetCrmServiceClient()
        {
            // In production, use secure configuration storage for credentials
            // This example uses dummy values - replace with your actual secure configuration method

            var connectionString = Environment.GetEnvironmentVariable("CRM_CONNECTION_STRING") ??
                @"Url=https://yourorg.crm.dynamics.com/;AuthType=OAuth;
                Username=user@yourorg.onmicrosoft.com;Password=YourSecurePassword123;
                ClientId=YourClientId;
                RedirectUri=app://YourRedirectAppId;
                LoginPrompt=Auto;RequireNewInstance=True";

            var service = new CrmServiceClient(connectionString);

            if (service.LastCrmError.Contains("Authentication"))
            {
                throw new Exception("Authentication failed. Please check credentials.");
            }

            return service;
        }

        /// <summary>
        /// Fetches users with specific segment, LOB and role filters
        /// </summary>
        private static List<UserData> FetchFilteredUsers(IOrganizationService service)
        {
            try
            {
                var query = new QueryExpression("systemuser")
                {
                    ColumnSet = new ColumnSet(
                        "systemuserid",
                        "fullname",
                        "zox_segment",
                        "zox_lob",
                        "zox_role",
                        "internalemailaddress",
                        "parentsystemuserid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("zox_segment", ConditionOperator.Equal, 100000002),
                            new ConditionExpression("zox_lob", ConditionOperator.Equal, 100000000),
                            new ConditionExpression("zox_role", ConditionOperator.In,
                                new object[] { 515140004, 515140005, 100000006 })
                        }
                    }
                };

                var results = service.RetrieveMultiple(query);
                Console.WriteLine($"Found {results.Entities.Count} filtered users");

                return results.Entities.Select(user => new UserData
                {
                    UserId = user.Id,
                    FullName = user.GetAttributeValue<string>("fullname") ?? "N/A",
                    Email = user.GetAttributeValue<string>("internalemailaddress") ?? string.Empty,
                    Segment = user.GetAttributeValue<OptionSetValue>("zox_segment")?.Value ?? -1,
                    LOB = user.GetAttributeValue<OptionSetValue>("zox_lob")?.Value ?? -1,
                    Role = user.GetAttributeValue<OptionSetValue>("zox_role")?.Value ?? -1,
                    ManagerId = user.GetAttributeValue<EntityReference>("parentsystemuserid")?.Id ?? Guid.Empty
                }).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching users: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Processes user hierarchy starting with HPR users
        /// </summary>
        private static void ProcessUserHierarchy(IOrganizationService service, List<UserData> usersList)
        {
            var processedUsers = new HashSet<Guid>();

            // Start with HPR users (role code 515140005)
            var hprUsers = usersList.Where(u => u.Role == 515140005).ToList();

            Console.WriteLine($"Processing {hprUsers.Count} HPR users");

            foreach (var hprUser in hprUsers)
            {
                ProcessUserAndManagers(service, usersList, hprUser, processedUsers);
            }
        }

        /// <summary>
        /// Processes a user and their management hierarchy
        /// </summary>
        private static void ProcessUserAndManagers(IOrganizationService service, List<UserData> usersList,
            UserData user, HashSet<Guid> processedUsers)
        {
            if (processedUsers.Contains(user.UserId)) return;
            processedUsers.Add(user.UserId);

            try
            {
                // Clear caches for each user to prevent memory bloat
                _lookupCache.Clear();
                _geographyCache.Clear();

                // Get products for user and their subordinates
                var userProducts = GetOpportunityProductsForUserAndSubordinates(service, usersList, user, processedUsers);
                if (userProducts.Count == 0) return;

                // Generate and upload report
                GenerateAndSaveExcelToDataverse(service, userProducts, $"{user.FullName}'s Team", user.UserId);

                // Process up the management chain
                ProcessManagementChain(service, usersList, user, processedUsers, userProducts);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing user {user.FullName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Processes the management chain upwards from the current user
        /// </summary>
        private static void ProcessManagementChain(IOrganizationService service, List<UserData> usersList,
            UserData user, HashSet<Guid> processedUsers, List<OpportunityProductData> allProducts)
        {
            var currentManagerId = user.ManagerId;

            while (currentManagerId != Guid.Empty && !processedUsers.Contains(currentManagerId))
            {
                var manager = usersList.FirstOrDefault(u => u.UserId == currentManagerId);
                if (manager == null) break;

                processedUsers.Add(currentManagerId);

                // Get products for manager and their subordinates
                var managerProducts = GetOpportunityProductsForUserAndSubordinates(service, usersList, manager, processedUsers);
                allProducts = allProducts.Union(managerProducts, new OpportunityProductDataEqualityComparer()).ToList();

                // Generate and upload report for manager
                GenerateAndSaveExcelToDataverse(service, allProducts, $"{manager.FullName}'s Team", manager.UserId);

                currentManagerId = manager.ManagerId;
            }
        }

        /// <summary>
        /// Gets opportunity products for a user and their subordinates
        /// </summary>
        private static List<OpportunityProductData> GetOpportunityProductsForUserAndSubordinates(
            IOrganizationService service, List<UserData> usersList, UserData user, HashSet<Guid> processedUsers)
        {
            var allProducts = new List<OpportunityProductData>();

            // Get products for the current user
            var userProducts = GetOpportunityProductsForUser(service, user.UserId, user.LOB);
            allProducts.AddRange(userProducts);

            // Process subordinates
            ProcessSubordinates(service, usersList, user.UserId, processedUsers, allProducts);

            return allProducts.Distinct(new OpportunityProductDataEqualityComparer()).ToList();
        }

        /// <summary>
        /// Processes subordinates for a manager
        /// </summary>
        private static void ProcessSubordinates(IOrganizationService service, List<UserData> usersList,
            Guid managerId, HashSet<Guid> processedUsers, List<OpportunityProductData> allProducts)
        {
            var subordinates = GetSubordinates(usersList, managerId);

            foreach (var subordinate in subordinates)
            {
                if (processedUsers.Contains(subordinate.UserId)) continue;

                var subordinateProducts = GetOpportunityProductsForUser(service, subordinate.UserId, subordinate.LOB);
                allProducts.AddRange(subordinateProducts);
            }
        }

        /// <summary>
        /// Gets users who report to the specified manager
        /// </summary>
        private static List<UserData> GetSubordinates(List<UserData> usersList, Guid managerId)
        {
            return usersList.Where(u => u.ManagerId == managerId).ToList();
        }

        /// <summary>
        /// Gets opportunity products shared with a specific user
        /// </summary>
        private static List<OpportunityProductData> GetOpportunityProductsForUser(
            IOrganizationService service, Guid userId, int userLob)
        {
            var products = new List<OpportunityProductData>();

            // First get shared product IDs
            var sharedProductIds = GetSharedProductIds(service, userId);

            if (sharedProductIds.Any())
            {
                // Then fetch full product details
                products.AddRange(FetchOpportunityProducts(service, CreateProductQuery(sharedProductIds, userLob)));
            }

            return products.Distinct(new OpportunityProductDataEqualityComparer()).ToList();
        }

        /// <summary>
        /// Gets IDs of opportunity products shared with the user
        /// </summary>
        private static List<Guid> GetSharedProductIds(IOrganizationService service, Guid userId)
        {
            var sharedQuery = new QueryExpression("principalobjectaccess")
            {
                ColumnSet = new ColumnSet("objectid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("principalid", ConditionOperator.Equal, userId),
                        new ConditionExpression("objecttypecode", ConditionOperator.Equal, "zox_opportunityproduct"),
                        new ConditionExpression("accessrightsmask", ConditionOperator.GreaterThan, 0)
                    }
                }
            };

            var sharedResults = service.RetrieveMultiple(sharedQuery);
            return sharedResults.Entities.Select(e => e.GetAttributeValue<Guid>("objectid")).ToList();
        }

        /// <summary>
        /// Creates query for fetching opportunity products
        /// </summary>
        private static QueryExpression CreateProductQuery(List<Guid> productIds, int lob)
        {
            return new QueryExpression("zox_opportunityproduct")
            {
                ColumnSet = new ColumnSet(
                    "zox_opportunityproductid", "zox_name", "ownerid", "createdby", "zox_lead", "zox_prelead",
                    "zox_lob", "zox_opportunity", "zox_product", "zox_productstatus", "createdon",
                    "zox_project_", "zox_contractor_", "zox_ponumber", "zox_sonumber", "zox_potential_"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("zox_opportunityproductid", ConditionOperator.In, productIds),
                        new ConditionExpression("zox_lob", ConditionOperator.Equal, lob)
                    }
                }
            };
        }

        /// <summary>
        /// Fetches opportunity products with geography data
        /// </summary>
        private static List<OpportunityProductData> FetchOpportunityProducts(
            IOrganizationService service, QueryExpression query)
        {
            var results = service.RetrieveMultiple(query);
            var products = new List<OpportunityProductData>();

            foreach (var entity in results.Entities)
            {
                ProcessOpportunityProductEntity(service, entity, products);
            }

            return products;
        }

        /// <summary>
        /// Processes a single opportunity product entity
        /// </summary>
        private static void ProcessOpportunityProductEntity(
            IOrganizationService service, Entity entity, List<OpportunityProductData> products)
        {
            var preLeadRef = entity.GetAttributeValue<EntityReference>("zox_prelead");
            var leadRef = entity.GetAttributeValue<EntityReference>("zox_lead");
            var opportunityRef = entity.GetAttributeValue<EntityReference>("zox_opportunity");
            var productRef = entity.GetAttributeValue<EntityReference>("zox_product");
            var potentialMoney = entity.GetAttributeValue<decimal>("zox_potential_");
            var createdByRef = entity.GetAttributeValue<EntityReference>("createdby");

            var geographies = GetUniqueGeographies(service, preLeadRef, leadRef, opportunityRef);
            if (geographies.Count == 0) geographies.Add(string.Empty);

            foreach (var geography in geographies)
            {
                if (geography != string.Empty || geographies.Count == 1)
                {
                    products.Add(CreateOpportunityProductData(
                        service, entity, preLeadRef, leadRef, opportunityRef,
                        productRef, potentialMoney, createdByRef, geography));
                }
            }
        }

        /// <summary>
        /// Creates opportunity product data object from entity
        /// </summary>
        private static OpportunityProductData CreateOpportunityProductData(
            IOrganizationService service, Entity entity, EntityReference preLeadRef,
            EntityReference leadRef, EntityReference opportunityRef, EntityReference productRef,
            decimal potentialMoney, EntityReference createdByRef, string geography)
        {
            return new OpportunityProductData
            {
                Id = $"{entity.Id}|{geography}",
                ProductName = entity.GetAttributeValue<string>("zox_name") ?? string.Empty,
                OwnerId = createdByRef?.Id ?? Guid.Empty,
                Lead = GetCachedLookupName(service, "lead", leadRef, "fullname"),
                PreLead = GetCachedLookupName(service, "zox_prelead", preLeadRef, "zox_name"),
                Opportunity = GetCachedLookupName(service, "opportunity", opportunityRef, "name"),
                Product = GetCachedLookupName(service, "zox_productcode", productRef, "zox_name"),
                LOB = GetOptionSetText(service, "zox_opportunityproduct", "zox_lob",
                    entity.GetAttributeValue<OptionSetValue>("zox_lob")?.Value),
                CreatedOn = entity.GetAttributeValue<DateTime>("createdon"),
                Project = GetCachedLookupName(service, "zox_project",
                    entity.GetAttributeValue<EntityReference>("zox_project_"), "zox_name"),
                Contractor = GetCachedLookupName(service, "account",
                    entity.GetAttributeValue<EntityReference>("zox_contractor_"), "name"),
                PONumber = GetCachedLookupName(service, "zox_purchaseorder",
                    entity.GetAttributeValue<EntityReference>("zox_ponumber"), "zox_name"),
                SONumber = GetCachedLookupName(service, "salesorder",
                    entity.GetAttributeValue<EntityReference>("zox_sonumber"), "name"),
                Potential = Math.Round(potentialMoney).ToString(),
                LeadGeography = geography,
                Status = GetOptionSetText(service, "zox_opportunityproduct", "zox_productstatus",
                    entity.GetAttributeValue<OptionSetValue>("zox_productstatus")?.Value)
            };
        }

        /// <summary>
        /// Gets unique geographies from related records
        /// </summary>
        private static List<string> GetUniqueGeographies(
            IOrganizationService service, EntityReference preLeadRef,
            EntityReference leadRef, EntityReference opportunityRef)
        {
            var geographies = new HashSet<string>();

            if (preLeadRef != null) geographies.UnionWith(GetCachedGeographies(service, "zox_prelead", preLeadRef.Id));
            if (leadRef != null) geographies.UnionWith(GetCachedGeographies(service, "zox_lead", leadRef.Id));
            if (opportunityRef != null) geographies.UnionWith(GetCachedGeographies(service, "zox_opportunity", opportunityRef.Id));

            return geographies.Count > 0 ? geographies.ToList() : new List<string> { string.Empty };
        }

        /// <summary>
        /// Gets cached geography data for an entity
        /// </summary>
        private static List<string> GetCachedGeographies(
            IOrganizationService service, string entityName, Guid entityId)
        {
            if (_geographyCache.TryGetValue(entityId, out var cachedGeos)) return cachedGeos;

            var geographies = new List<string>();

            try
            {
                var query = new QueryExpression("zox_leadgeographymapping")
                {
                    ColumnSet = new ColumnSet("zox_region"),
                    Criteria = new FilterExpression
                    {
                        Conditions = { new ConditionExpression(entityName, ConditionOperator.Equal, entityId) }
                    }
                };

                var results = service.RetrieveMultiple(query);
                geographies.AddRange(
                    from entity in results.Entities
                    let geographyRef = entity.GetAttributeValue<EntityReference>("zox_region")
                    where geographyRef != null
                    select GetCachedLookupName(service, "zox_regionmaster", geographyRef, "zox_name")
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching geographies: {ex.Message}");
            }

            _geographyCache[entityId] = geographies.Count > 0 ? geographies : new List<string> { string.Empty };
            return _geographyCache[entityId];
        }

        /// <summary>
        /// Gets a lookup value from cache or CRM
        /// </summary>
        private static string GetCachedLookupName(
            IOrganizationService service, string entityName,
            EntityReference lookupRef, string fieldName)
        {
            if (lookupRef == null) return string.Empty;
            if (_lookupCache.TryGetValue(lookupRef.Id, out var cachedName)) return cachedName;

            try
            {
                var entity = service.Retrieve(entityName, lookupRef.Id, new ColumnSet(fieldName));
                var name = entity.GetAttributeValue<string>(fieldName) ?? string.Empty;
                _lookupCache[lookupRef.Id] = name;
                return name;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving {entityName} {lookupRef.Id}: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the text label for an option set value
        /// </summary>
        private static string GetOptionSetText(
            IOrganizationService service, string entityName,
            string attributeName, int? optionSetValue)
        {
            if (!optionSetValue.HasValue) return "Open";

            try
            {
                OptionSetMetadata optionSetMetadata;

                if (attributeName == "zox_productstatus")
                {
                    // Handle local option set
                    var attributeRequest = new RetrieveAttributeRequest
                    {
                        EntityLogicalName = entityName,
                        LogicalName = attributeName
                    };
                    var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
                    var attributeMetadata = (PicklistAttributeMetadata)attributeResponse.AttributeMetadata;
                    optionSetMetadata = attributeMetadata.OptionSet;
                }
                else
                {
                    // Handle global option set
                    var optionSetRequest = new RetrieveOptionSetRequest { Name = attributeName };
                    var optionSetResponse = (RetrieveOptionSetResponse)service.Execute(optionSetRequest);
                    optionSetMetadata = (OptionSetMetadata)optionSetResponse.OptionSetMetadata;
                }

                var option = optionSetMetadata.Options.FirstOrDefault(o => o.Value == optionSetValue.Value);
                return option?.Label.UserLocalizedLabel?.Label ?? string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving option set {attributeName}: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Generates Excel report and saves to Dataverse
        /// </summary>
        private static void GenerateAndSaveExcelToDataverse(
            IOrganizationService service, List<OpportunityProductData> products,
            string reportName, Guid userId)
        {
            string filePath = string.Empty;

            try
            {
                reportName = string.Join("_", reportName.Split(Path.GetInvalidFileNameChars()));
                filePath = Path.Combine(Path.GetTempPath(), $"{reportName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Opportunity Products");

                    // Set up report headers
                    SetupReportHeaders(worksheet);

                    // Add data rows
                    PopulateReportData(service, worksheet, products);

                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(filePath);
                }

                // Upload to Dataverse
                UploadExcelToDataverse(service, File.ReadAllBytes(filePath), reportName, userId);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating report: {ex.Message}");
            }
            finally
            {
                // Clean up temp file
                if (File.Exists(filePath))
                {
                    try { File.Delete(filePath); }
                    catch { /* Ignore cleanup errors */ }
                }
            }
        }

        /// <summary>
        /// Sets up the report headers and formatting
        /// </summary>
        private static void SetupReportHeaders(IXLWorksheet worksheet)
        {
            var headers = new[]
            {
                "S.No.", "Pre Lead", "Lead", "Lob", "Opportunity", "Project",
                "Created By", "Created On", "Product", "Potential", "Contractor",
                "PO Number", "SO Number", "Lead Geography", "OP Aging", "Status"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            }
        }

        /// <summary>
        /// Populates the report with data rows
        /// </summary>
        private static void PopulateReportData(
            IOrganizationService service, IXLWorksheet worksheet,
            List<OpportunityProductData> products)
        {
            for (int row = 0; row < products.Count; row++)
            {
                var product = products[row];
                int col = 1;

                worksheet.Cell(row + 2, col++).Value = row + 1;
                worksheet.Cell(row + 2, col++).Value = product.PreLead;
                worksheet.Cell(row + 2, col++).Value = product.Lead;
                worksheet.Cell(row + 2, col++).Value = product.LOB;
                worksheet.Cell(row + 2, col++).Value = product.Opportunity;
                worksheet.Cell(row + 2, col++).Value = product.Project;
                worksheet.Cell(row + 2, col++).Value = GetCachedLookupName(
                    service, "systemuser", new EntityReference("systemuser", product.OwnerId), "fullname");
                worksheet.Cell(row + 2, col++).Value = product.CreatedOn.ToString("yyyy-MM-dd");
                worksheet.Cell(row + 2, col++).Value = product.Product;
                worksheet.Cell(row + 2, col++).Value = product.Potential;
                worksheet.Cell(row + 2, col++).Value = product.Contractor;
                worksheet.Cell(row + 2, col++).Value = product.PONumber;
                worksheet.Cell(row + 2, col++).Value = product.SONumber;
                worksheet.Cell(row + 2, col++).Value = product.LeadGeography;
                worksheet.Cell(row + 2, col++).Value = (DateTime.Now - product.CreatedOn).Days;
                worksheet.Cell(row + 2, col++).Value = product.Status;
            }
        }

        /// <summary>
        /// Uploads Excel file to Dataverse using chunked upload
        /// </summary>
        private static void UploadExcelToDataverse(
            IOrganizationService service, byte[] fileContent,
            string reportName, Guid userId)
        {
            try
            {
                // Initialize upload
                var initializeRequest = new InitializeFileBlocksUploadRequest
                {
                    Target = new EntityReference("systemuser", userId),
                    FileAttributeName = "zx_file",
                    FileName = $"{reportName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx"
                };

                var initializeResponse = (InitializeFileBlocksUploadResponse)service.Execute(initializeRequest);
                var blockList = new List<string>();
                int blockSize = 4 * 1024 * 1024; // 4MB chunks

                // Upload file in chunks
                for (int i = 0; i < (fileContent.Length + blockSize - 1) / blockSize; i++)
                {
                    byte[] buffer = fileContent.Skip(i * blockSize).Take(blockSize).ToArray();
                    var blockId = Convert.ToBase64String(Encoding.UTF8.GetBytes(Guid.NewGuid().ToString()));

                    service.Execute(new UploadBlockRequest
                    {
                        BlockData = buffer,
                        BlockId = blockId,
                        FileContinuationToken = initializeResponse.FileContinuationToken
                    });

                    blockList.Add(blockId);
                }

                // Commit the upload
                service.Execute(new CommitFileBlocksUploadRequest
                {
                    BlockList = blockList.ToArray(),
                    FileContinuationToken = initializeResponse.FileContinuationToken,
                    FileName = initializeRequest.FileName,
                    MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                });

                Console.WriteLine($"Successfully uploaded report for user {userId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uploading file to Dataverse: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Represents user data from CRM
    /// </summary>
    public class UserData
    {
        public Guid UserId { get; set; }
        public string FullName { get; set; }
        public string Email { get; set; }
        public int Segment { get; set; }
        public int LOB { get; set; }
        public int Role { get; set; }
        public Guid ManagerId { get; set; }
    }

    /// <summary>
    /// Represents opportunity product data from CRM
    /// </summary>
    public class OpportunityProductData
    {
        public string Id { get; set; }
        public string ProductName { get; set; }
        public Guid OwnerId { get; set; }
        public string Lead { get; set; }
        public string PreLead { get; set; }
        public string Opportunity { get; set; }
        public string Product { get; set; }
        public string LOB { get; set; }
        public DateTime CreatedOn { get; set; }
        public string Project { get; set; }
        public string Contractor { get; set; }
        public string PONumber { get; set; }
        public string SONumber { get; set; }
        public string Potential { get; set; }
        public string LeadGeography { get; set; }
        public string Status { get; set; }
    }

    /// <summary>
    /// Equality comparer for OpportunityProductData
    /// </summary>
    public class OpportunityProductDataEqualityComparer : IEqualityComparer<OpportunityProductData>
    {
        public bool Equals(OpportunityProductData x, OpportunityProductData y) => x?.Id == y?.Id;
        public int GetHashCode(OpportunityProductData obj) => obj.Id.GetHashCode();
    }
}