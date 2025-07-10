using log4net;
using log4net.Config;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
namespace CrmConnectionTool
{
    class Program
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));

        static IOrganizationService ConnectToDynamics(string connectionString)
        {
            CrmServiceClient serviceClient = new CrmServiceClient(connectionString);
            if (!serviceClient.IsReady)
            {
                throw new Exception("Failed to connect to Dynamics 365");
            }
            return serviceClient;
        }

        static string GetFormXml(IOrganizationService service, string entityName, string formName)
        {
            QueryExpression query = new QueryExpression("systemform")
            {
                ColumnSet = new ColumnSet("formxml"),
                Criteria = new FilterExpression()
            };
            query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, RetrieveEntityObjectTypeCode(service, entityName));
            query.Criteria.AddCondition("name", ConditionOperator.Equal, formName);

            return service.RetrieveMultiple(query).Entities?.FirstOrDefault()?["formxml"]?.ToString() ?? throw new Exception($"Form '{formName}' not found for entity '{entityName}'");
        }

        static int RetrieveEntityObjectTypeCode(IOrganizationService service, string entityName)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest
            {
                LogicalName = entityName,
                EntityFilters = EntityFilters.Entity
            };
            RetrieveEntityResponse response = (RetrieveEntityResponse)service.Execute(request);
            return response.EntityMetadata.ObjectTypeCode ?? throw new Exception($"ObjectTypeCode not found for entity '{entityName}'");
        }

        static string GetLabel(XmlNode node)
        {
            XmlNode labelNode = node.SelectSingleNode(".//labels/label[@description]");
            if (labelNode != null && labelNode.Attributes["description"] != null)
            {
                return labelNode.Attributes["description"].Value;
            }
            return "Unknown Label";
        }

        // Structure pour stocker toutes les informations nécessaires sur les éléments du formulaire
        class FormElement
        {
            public string Id { get; set; }
            public string DisplayName { get; set; }
            public string FieldName { get; set; }  // Nom logique du champ
            public string CustomLabel { get; set; } // Étiquette personnalisée sur le formulaire
            public bool HasCustomLabel { get; set; } // Indique si le champ a une étiquette personnalisée
            public string TabId { get; set; }
            public string TabName { get; set; }
            public string SectionId { get; set; }
            public string SectionName { get; set; }
            public bool Visible { get; set; } // Indique si le champ est visible ou masqué
            public bool Required { get; set; } // Indique si le champ est obligatoire
            public string ControlType { get; set; } // Type de contrôle (texte, lookup, etc.)
        }
        // Structure pour stocker les informations sur les labels
        class LabelElement
        {
            public string Id { get; set; }
            public string Type { get; set; }
            public string Label { get; set; }
        }

        // Méthode pour extraire les champs du formulaire avec leurs onglets et sections correspondants
        // et détection des étiquettes personnalisées
        static List<XmlNode> ExtractRowsFromFormXml(string formXml)
        {
            if (string.IsNullOrEmpty(formXml))
                return new List<XmlNode>();

            var fields = new List<FormElement>();
            var rows = new List<XmlNode>();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(formXml);

            // Créer un dictionnaire pour stocker les noms des onglets et sections
            var tabNames = new Dictionary<string, string>();
            var sectionNames = new Dictionary<string, string>();

            // Extraire les noms des onglets
            XmlNodeList tabNodes = doc.SelectNodes("//tab");
            foreach (XmlNode tabNode in tabNodes)
            {
                string tabId = tabNode.Attributes?["id"]?.Value;
                if (!string.IsNullOrEmpty(tabId))
                {
                    string tabName = GetLabel(tabNode);
                    tabNames[tabId] = tabName;

                    // Extraire les noms des sections pour cet onglet
                    XmlNodeList sectionNodes = tabNode.SelectNodes(".//section");
                    foreach (XmlNode sectionNode in sectionNodes)
                    {
                        string sectionId = sectionNode.Attributes?["id"]?.Value;
                        if (!string.IsNullOrEmpty(sectionId))
                        {
                            string sectionName = GetLabel(sectionNode);
                            sectionNames[sectionId] = sectionName;
                        }
                    }
                }
            }

            // Extraire les champs et leurs relations avec les onglets et sections
            XmlNodeList fieldNodes = doc.SelectNodes("//control[@datafieldname]");
            foreach (XmlNode fieldNode in fieldNodes)
            {
                try
                {
                    // Obtenez l'ID du champ et son nom logique
                    string fieldId = fieldNode.Attributes?["id"]?.Value ?? "Unknown Field";
                    string fieldName = fieldNode.Attributes?["datafieldname"]?.Value ?? fieldId;

                    // Détermine si le champ est visible
                    bool visible = true;
                    XmlNode visibleNode = fieldNode.SelectSingleNode(".//visible");
                    if (visibleNode != null)
                    {
                        visible = visibleNode.InnerText.ToLower() != "false";
                    }

                    // Détermine si le champ est obligatoire
                    bool required = false;
                    XmlNode requiredNode = fieldNode.SelectSingleNode(".//required");
                    if (requiredNode != null)
                    {
                        required = requiredNode.InnerText.ToLower() == "true";
                    }

                    // Détermine le type de contrôle
                    string controlType = fieldNode.Attributes?["classid"]?.Value ?? "Unknown";

                    // Vérifier si le contrôle a une étiquette personnalisée
                    string customLabel = null;
                    bool hasCustomLabel = false;

                    // Approche exhaustive pour trouver les étiquettes personnalisées
                    // 1. Recherche dans les labels directs
                    XmlNode labelNode = fieldNode.SelectSingleNode("./labels/label");
                    if (labelNode != null && labelNode.Attributes?["description"] != null)
                    {
                        customLabel = labelNode.Attributes["description"].Value;
                        hasCustomLabel = true;
                    }

                    // 2. Recherche dans cellLabels
                    if (!hasCustomLabel)
                    {
                        XmlNode cellLabelsNode = fieldNode.SelectSingleNode(".//cellLabels/label");
                        if (cellLabelsNode != null && cellLabelsNode.Attributes?["description"] != null)
                        {
                            customLabel = cellLabelsNode.Attributes["description"].Value;
                            hasCustomLabel = true;
                        }
                    }

                    // 3. Recherche dans l'attribut label
                    if (!hasCustomLabel && fieldNode.Attributes?["label"] != null)
                    {
                        customLabel = fieldNode.Attributes["label"].Value;
                        hasCustomLabel = true;
                    }

                    // 4. Recherche dans customlabel
                    if (!hasCustomLabel)
                    {
                        XmlNode customLabelNode = fieldNode.SelectSingleNode(".//customlabel");
                        if (customLabelNode != null && !string.IsNullOrEmpty(customLabelNode.InnerText))
                        {
                            customLabel = customLabelNode.InnerText;
                            hasCustomLabel = true;
                        }
                    }

                    // 5. Recherche dans les paramètres de cellule
                    if (!hasCustomLabel)
                    {
                        XmlNodeList cellParams = fieldNode.SelectNodes(".//cellLabelAttributes");
                        foreach (XmlNode cellParam in cellParams)
                        {
                            if (cellParam?.InnerText?.Contains("customlabel") == true)
                            {
                                try
                                {
                                    XmlDocument paramDoc = new XmlDocument();
                                    paramDoc.LoadXml(cellParam.InnerText);
                                    XmlNode labelTextNode = paramDoc.SelectSingleNode("//customlabel");
                                    if (labelTextNode != null && !string.IsNullOrEmpty(labelTextNode.InnerText))
                                    {
                                        customLabel = labelTextNode.InnerText;
                                        hasCustomLabel = true;
                                        break;
                                    }
                                }
                                catch
                                {
                                    // Ignorer les erreurs de parsing
                                }
                            }
                        }
                    }

                    // Parcourez les parents pour trouver la section et l'onglet
                    XmlNode currentNode = fieldNode;
                    string sectionId = "Unknown Section";
                    string tabId = "Unknown Tab";

                    // Cherchez la section parente
                    while (currentNode != null && currentNode.Name != "section")
                    {
                        currentNode = currentNode.ParentNode;
                    }

                    if (currentNode != null)
                    {
                        sectionId = currentNode.Attributes?["id"]?.Value ?? "Unknown Section";

                        // Continuez à chercher l'onglet parent
                        while (currentNode != null && currentNode.Name != "tab")
                        {
                            currentNode = currentNode.ParentNode;
                        }

                        if (currentNode != null)
                        {
                            tabId = currentNode.Attributes?["id"]?.Value ?? "Unknown Tab";
                        }
                    }

                    // Obtenez les noms des onglets et sections
                    string tabName = tabNames.ContainsKey(tabId) ? tabNames[tabId] : "Unknown Tab";
                    string sectionName = sectionNames.ContainsKey(sectionId) ? sectionNames[sectionId] : "Unknown Section";

                    fields.Add(new FormElement
                    {
                        Id = fieldId,
                        DisplayName = fieldName,
                        FieldName = fieldName,
                        CustomLabel = customLabel,
                        HasCustomLabel = hasCustomLabel,
                        TabId = tabId,
                        TabName = tabName,
                        SectionId = sectionId,
                        SectionName = sectionName,
                        Visible = visible,
                        Required = required,
                        ControlType = controlType
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing field: {ex.Message}");
                }
            }

            XmlNodeList rowNodes = doc.SelectNodes("//row");
            foreach (XmlNode rowNode in rowNodes)
            {
                rows.Add(rowNode);
            }

            return rows;
        }
        //// Méthode pour extraire les champs du formulaire avec leurs onglets et sections correspondants
        static List<FormElement> ExtractFieldsFromFormXml(string formXml)
        {
            var fields = new List<FormElement>();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(formXml);

            // Créer un dictionnaire pour stocker les noms des onglets et sections
            var tabNames = new Dictionary<string, string>();
            var sectionNames = new Dictionary<string, string>();

            // Extraire les noms des onglets
            XmlNodeList tabNodes = doc.SelectNodes("//tab");
            foreach (XmlNode tabNode in tabNodes)
            {
                string tabId = tabNode.Attributes?["id"]?.Value;
                if (!string.IsNullOrEmpty(tabId))
                {
                    string tabName = GetLabel(tabNode);
                    tabNames[tabId] = tabName;

                    // Extraire les noms des sections pour cet onglet
                    XmlNodeList sectionNodes = tabNode.SelectNodes(".//section");
                    foreach (XmlNode sectionNode in sectionNodes)
                    {
                        string sectionId = sectionNode.Attributes?["id"]?.Value;
                        if (!string.IsNullOrEmpty(sectionId))
                        {
                            string sectionName = GetLabel(sectionNode);
                            sectionNames[sectionId] = sectionName;
                        }
                    }
                }
            }

            // Extraire les champs et leurs relations avec les onglets et sections
            XmlNodeList fieldNodes = doc.SelectNodes("//control[@datafieldname]");
            foreach (XmlNode fieldNode in fieldNodes)
            {
                // Obtenez l'ID du champ et son nom logique
                string fieldId = fieldNode.Attributes?["id"]?.Value ?? "Unknown Field";
                string fieldName = fieldNode.Attributes?["datafieldname"]?.Value ?? fieldId;

                // Vérifier si le contrôle a une étiquette personnalisée
                string customLabel = null;
                bool hasCustomLabel = false;

                // Chercher les étiquettes personnalisées dans le nœud 'label' du contrôle
                // Première méthode : chercher un nœud label directement sous le contrôle
                XmlNode labelNode = fieldNode.SelectSingleNode(".//labels/label[@description]");
                if (labelNode != null && labelNode.Attributes["description"] != null)
                {
                    customLabel = labelNode.Attributes["description"].Value;
                    hasCustomLabel = true;
                }

                // Deuxième méthode : rechercher un nœud de paramètre spécifique qui définit l'étiquette
                XmlNode paramNode = fieldNode.SelectSingleNode(".//customlabel");
                if (paramNode != null && paramNode.InnerText != null)
                {
                    customLabel = paramNode.InnerText;
                    hasCustomLabel = true;
                }

                // Troisième méthode : rechercher dans l'attribut 'label' s'il existe
                if (fieldNode.Attributes?["label"] != null)
                {
                    customLabel = fieldNode.Attributes["label"].Value;
                    hasCustomLabel = true;
                }

                // Quatrième méthode : rechercher dans l'élément 'labelid'
                XmlNode labelIdNode = fieldNode.SelectSingleNode(".//labelid");
                if (labelIdNode != null && !string.IsNullOrEmpty(labelIdNode.InnerText))
                {
                    // Trouver l'élément correspondant dans la section des labels du formulaire
                    XmlNode formLabelNode = doc.SelectSingleNode($"//label[@id='{labelIdNode.InnerText}']");
                    if (formLabelNode != null && formLabelNode.Attributes?["description"] != null)
                    {
                        customLabel = formLabelNode.Attributes["description"].Value;
                        hasCustomLabel = true;
                    }
                }

                // Parcourez les parents pour trouver la section et l'onglet
                XmlNode currentNode = fieldNode;
                string sectionId = "Unknown Section";
                string tabId = "Unknown Tab";

                // Cherchez la section parente
                while (currentNode != null && currentNode.Name != "section")
                {
                    currentNode = currentNode.ParentNode;
                }

                if (currentNode != null)
                {
                    sectionId = currentNode.Attributes?["id"]?.Value ?? "Unknown Section";

                    // Continuez à chercher l'onglet parent
                    while (currentNode != null && currentNode.Name != "tab")
                    {
                        currentNode = currentNode.ParentNode;
                    }

                    if (currentNode != null)
                    {
                        tabId = currentNode.Attributes?["id"]?.Value ?? "Unknown Tab";
                    }
                }

                // Obtenez les noms des onglets et sections
                string tabName = tabNames.ContainsKey(tabId) ? tabNames[tabId] : "Unknown Tab";
                string sectionName = sectionNames.ContainsKey(sectionId) ? sectionNames[sectionId] : "Unknown Section";

                fields.Add(new FormElement
                {
                    Id = fieldId,
                    DisplayName = fieldName,
                    FieldName = fieldName,
                    CustomLabel = customLabel,
                    HasCustomLabel = hasCustomLabel,
                    TabId = tabId,
                    TabName = tabName,
                    SectionId = sectionId,
                    SectionName = sectionName
                });
            }

            return fields;
        }

        // Méthode pour extraire les labels des onglets et sections
        static List<LabelElement> ExtractLabelsFromFormXml(string formXml)
        {
            var elements = new List<LabelElement>();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(formXml);

            XmlNodeList tabs = doc.SelectNodes("//tab");
            foreach (XmlNode tab in tabs)
            {
                string tabId = tab.Attributes?["id"]?.Value ?? "Unknown GUID";
                string tabLabel = GetLabel(tab);
                elements.Add(new LabelElement { Id = tabId, Type = "Tab", Label = tabLabel });

                XmlNodeList sections = tab.SelectNodes(".//section");
                foreach (XmlNode section in sections)
                {
                    string sectionId = section.Attributes?["id"]?.Value ?? "Unknown GUID";
                    string sectionLabel = GetLabel(section);
                    elements.Add(new LabelElement { Id = sectionId, Type = "Section", Label = sectionLabel });
                }
            }
            return elements;
        }
        static Dictionary<string, Dictionary<string, string>> ExtractFieldTranslations(string formXml)
        {
            var result = new Dictionary<string, Dictionary<string, string>>();
            var doc = new XmlDocument();
            doc.LoadXml(formXml);
            foreach (XmlNode ctrl in doc.SelectNodes("//control[@datafieldname]"))
            {
                var name = ctrl.Attributes?["datafieldname"]?.Value;
                if (string.IsNullOrEmpty(name)) continue;
                var labels = new Dictionary<string, string>();
                foreach (XmlNode label in ctrl.SelectNodes(".//labels/label"))
                {
                    string lang = label.Attributes?["languagecode"]?.Value ?? "default";
                    string desc = label.Attributes?["description"]?.Value ?? string.Empty;
                    labels[lang] = desc;
                }
                if (labels.Count > 0)
                    result[name] = labels;
            }
            return result;
        }


        // Méthode pour comparer les labels entre deux formulaires
        static void CompareLabels(List<LabelElement> labels1, List<LabelElement> labels2)
        {
            Console.WriteLine("\n--- Changed Labels ---");
            foreach (var label1 in labels1)
            {
                var label2 = labels2.FirstOrDefault(l => l.Id == label1.Id && l.Type == label1.Type);
                if (label2 != null && label1.Label != label2.Label)
                {
                    Console.WriteLine($"Label changed in {label1.Type}: '{label1.Label}' -> '{label2.Label}'");
                }
            }
        }

        static void PrintRowDetails(XmlNode row)
        {
            XmlNode control = row.SelectSingleNode(".//control");
            XmlNode cell = row.SelectSingleNode(".//cell");
            XmlNode tab = row.SelectSingleNode("ancestor::tab");
            XmlNode section = row.SelectSingleNode("ancestor::section");
            XmlNode labelsNode = row.SelectSingleNode(".//labels");

            if (control == null)
            {
                Console.WriteLine("  Unable to find row details.");
                return;
            }

            string fieldName = control.Attributes["datafieldname"]?.Value ?? "Unknown";
            string tabName = tab?.Attributes["name"]?.Value ?? "Unknown Tab";
            string sectionName = section?.Attributes["name"]?.Value ?? "Unknown Section";
            string cellId = cell?.Attributes["id"]?.Value ?? "Unknown ID";

            Console.WriteLine($"  Tab: {tabName} - Section: {sectionName} - Field: {fieldName} - ID: {cellId}");

            // Afficher les labels
            Console.WriteLine("  Labels:");
            if (labelsNode != null)
            {
                XmlNodeList labelNodes = labelsNode.SelectNodes("./label");
                if (labelNodes != null && labelNodes.Count > 0)
                {
                    foreach (XmlNode labelNode in labelNodes)
                    {
                        string langCode = labelNode.Attributes?["languagecode"]?.Value ?? "N/A";
                        string description = labelNode.Attributes?["description"]?.Value ?? "Empty";
                        string labelText = labelNode.Attributes?["text"]?.Value ?? "N/A";

                        Console.WriteLine($"Lang: {langCode} - Desc: {description} - Text: {labelText}");
                    }
                }
                else
                {
                    Console.WriteLine("    Label not found");
                }
            }
            else
            {
                Console.WriteLine("    No labels node");
            }
        }

        static void CompareRowStructures(XmlNode row1, XmlNode row2, string cellId)
        {
            // Comparer les attributs des cellules
            XmlNode cell1 = row1.SelectSingleNode(".//cell");
            XmlNode cell2 = row2.SelectSingleNode(".//cell");

            if (cell1 == null || cell2 == null)
                return;

            // Collecter tous les attributs
            HashSet<string> allAttributes = new HashSet<string>();
            foreach (XmlAttribute attr in cell1.Attributes)
                allAttributes.Add(attr.Name);
            foreach (XmlAttribute attr in cell2.Attributes)
                allAttributes.Add(attr.Name);

            bool differenceFound = false;
            StringBuilder differences = new StringBuilder();
            differences.AppendLine($"Différences pour la row avec ID de cellule '{cellId}':");

            // Comparer chaque attribut
            foreach (string attrName in allAttributes)
            {
                string value1 = cell1.Attributes[attrName]?.Value;
                string value2 = cell2.Attributes[attrName]?.Value;

                if (value1 != value2)
                {
                    differenceFound = true;
                    if (value1 == null)
                        differences.AppendLine($"  - Attribut '{attrName}' added in destination form: '{value2}'");
                    else if (value2 == null)
                        differences.AppendLine($"  - Attribut '{attrName}' removed in destination form (value was: '{value1}')");
                    else
                        differences.AppendLine($"  - Attribut '{attrName}' updated: '{value1}' -> '{value2}'");
                }
            }

            // Comparer les contrôles
            XmlNode control1 = row1.SelectSingleNode(".//control");
            XmlNode control2 = row2.SelectSingleNode(".//control");

            if (control1 != null && control2 != null)
            {
                // Comparer les attributs des contrôles
                allAttributes.Clear();
                foreach (XmlAttribute attr in control1.Attributes)
                    allAttributes.Add(attr.Name);
                foreach (XmlAttribute attr in control2.Attributes)
                    allAttributes.Add(attr.Name);

                foreach (string attrName in allAttributes)
                {
                    string value1 = control1.Attributes[attrName]?.Value;
                    string value2 = control2.Attributes[attrName]?.Value;

                    if (value1 != value2)
                    {
                        differenceFound = true;
                        if (value1 == null)
                            differences.AppendLine($"  - Attribut de contrôle '{attrName}' ajouté: '{value2}'");
                        else if (value2 == null)
                            differences.AppendLine($"  - Attribut de contrôle '{attrName}' supprimé (valeur était: '{value1}')");
                        else
                            differences.AppendLine($"  - Attribut de contrôle '{attrName}' modifié: '{value1}' -> '{value2}'");
                    }
                }
            }

            // Comparer les labels
            CompareLabelsInRow(row1, row2, cellId, differences);

            // Afficher les différences si trouvées
            if (differenceFound)
                Console.WriteLine(differences.ToString());
        }

        static void CompareLabelsInRow(XmlNode row1, XmlNode row2, string cellId, StringBuilder differences)
        {
            XmlNode labels1 = row1.SelectSingleNode(".//labels");
            XmlNode labels2 = row2.SelectSingleNode(".//labels");

            if (labels1 == null || labels2 == null)
                return;

            XmlNodeList labelNodes1 = labels1.SelectNodes("./label");
            XmlNodeList labelNodes2 = labels2.SelectNodes("./label");

            Dictionary<string, string> labelMap1 = new Dictionary<string, string>();
            Dictionary<string, string> labelMap2 = new Dictionary<string, string>();

            // Collecter les labels par code de langue
            foreach (XmlNode label in labelNodes1)
            {
                string langCode = label.Attributes?["languagecode"]?.Value ?? "default";
                string desc = label.Attributes?["description"]?.Value;
                if (desc != null)
                    labelMap1[langCode] = desc;
            }

            foreach (XmlNode label in labelNodes2)
            {
                string langCode = label.Attributes?["languagecode"]?.Value ?? "default";
                string desc = label.Attributes?["description"]?.Value;
                if (desc != null)
                    labelMap2[langCode] = desc;
            }

            // Comparer les labels
            HashSet<string> allLangCodes = new HashSet<string>();
            foreach (var code in labelMap1.Keys) allLangCodes.Add(code);
            foreach (var code in labelMap2.Keys) allLangCodes.Add(code);

            bool labelDifferenceFound = false;

            foreach (string code in allLangCodes)
            {
                labelMap1.TryGetValue(code, out string desc1);
                labelMap2.TryGetValue(code, out string desc2);

                if (desc1 != desc2)
                {
                    if (!labelDifferenceFound)
                    {
                        differences.AppendLine($"Différences de labels pour la row avec ID de cellule '{cellId}':");
                        labelDifferenceFound = true;
                    }

                    if (desc1 == null)
                        differences.AppendLine($"  - Label ajouté pour langue '{code}': '{desc2}'");
                    else if (desc2 == null)
                        differences.AppendLine($"  - Label supprimé pour langue '{code}' (valeur était: '{desc1}')");
                    else
                        differences.AppendLine($"  - Label modifié pour langue '{code}': '{desc1}' -> '{desc2}'");
                }
            }
        }

        static void Main(string[] args)
        {
            XmlConfigurator.Configure(new FileInfo("log4net.config"));
            var clientCredentialsUser = new ClientCredentials();
            //clientCredentialsUser.UserName.UserName = "s_prot_admin_u";
            //clientCredentialsUser.UserName.Password = "";
            //var lstOrg = ManageUser.GetOrganizations("https://prot6-zone1.uat.icrc.org/XRMServices/2011/Discovery.svc", "s_prot_admin_u", "ddexogin36", "gva");
            //foreach (var org in lstOrg) {
            //    var orgConString = new Uri($"https://prot6-zone1.uat.icrc.org/{org}/XRMServices/2011/Organization.svc");
            //    ManageUser.CreateUser("gva\\s_prom_mssql_u",org, orgConString, "System Customizer", clientCredentialsUser);
            //}

            var clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = "A053751";
            clientCredentials.UserName.Password = Settings.Default.Password;

            string conn1 = $"AuthType=IFD; Url=https://usa.flanswerstest.org/usa; Username=fespla@icrc.org; Password={Settings.Default.Password}; Domain=ICRC;"; // string conn1 = $"AuthType=AD; Url=https://usa.flanswerstest.org/; Domain=GVA; Username=a053751; Password=;";
            string conn2 = $"AuthType=IFD; Url=https://usaprodal.flanswerstest.org/usaprodal;Username=fespla@icrc.org; Password={Settings.Default.Password}; Domain=ICRC;";

            var service1 = ConnectToDynamics(conn1);
            var service2 = ConnectToDynamics(conn2);

            List<(string EntityName, string FormName)> formsToCompare = new List<(string, string)>
            {
                ("contact", "Information"),
                ("systemuser", "User"),
                ("core_activity", "Information"),
                ("core_accompanyingperson", "Information"),
                ("core_detactivitydetail", "Information"),
                ("core_history", "Information"),
                ("task", "Information"),
            };

            //CompareFormCustomization.CompareForm(service1,service2, formsToCompare);
            foreach (var (entity, form) in formsToCompare)
            {
                LogicalThreadContext.Properties["Entity"] = entity;
                LogicalThreadContext.Properties["Type"] = form;
                try
                {
                    Console.WriteLine($"\n==================================================");
                    Console.WriteLine($"Comparing form '{form}' for entity '{entity}':");
                    Console.WriteLine($"==================================================");

                    string formXml1 = GetFormXml(service1, entity, form);
                    string formXml2 = GetFormXml(service2, entity, form);

                    if (string.IsNullOrEmpty(formXml1) || string.IsNullOrEmpty(formXml2))
                    {
                        Console.WriteLine($"ERROR: One or both forms not found: '{form}' for entity '{entity}'");
                        continue;
                    }

                    // Vérification rapide pour voir si les formulaires sont identiques
                    if (formXml1 == formXml2)
                    {
                        log.Info("Form identical");
                        Console.WriteLine("The forms are identical. No changes detected.");
                        continue;
                    }

                    //get fields and rows
                    var fields1 = ExtractFieldsFromFormXml(formXml1); 
                    var rows1 = ExtractRowsFromFormXml(formXml1);
                    var fields2 = ExtractFieldsFromFormXml(formXml2);
                    var rows2 = ExtractRowsFromFormXml(formXml2);

                    // Créer des dictionnaires pour mapper les rows par leurs IDs de cellule
                    Dictionary<string, XmlNode> rowMap1 = new Dictionary<string, XmlNode>();
                    Dictionary<string, XmlNode> rowMap2 = new Dictionary<string, XmlNode>();

                    // Mapper les rows par l'ID de leur cellule
                    foreach (XmlNode row in rows1)
                    {
                        XmlNode cell = row.SelectSingleNode(".//cell");
                        if (cell != null && cell.Attributes["id"] != null)
                        {
                            rowMap1[cell.Attributes["id"].Value] = row;
                        }
                    }

                    foreach (XmlNode row in rows2)
                    {
                        XmlNode cell = row.SelectSingleNode(".//cell");
                        if (cell != null && cell.Attributes["id"] != null)
                        {
                            rowMap2[cell.Attributes["id"].Value] = row;
                        }
                    }

                    // COmpare rows with same key
                    Console.WriteLine("Comparing Rows:");
                    foreach (var kvp in rowMap1)
                    {
                        string key = kvp.Key;
                        XmlNode row1 = kvp.Value;

                        if (rowMap2.TryGetValue(key, out XmlNode row2))
                        {
                            // Compare attributs and components
                            CompareRowStructures(row1, row2, key);
                        }
                        else
                        {
                            Console.WriteLine($"Row non present in destination:");
                            PrintRowDetails(row1);
                        }
                    }

                    // Identifier les rows présentes uniquement dans le deuxième formulaire
                    foreach (var kvp in rowMap2)
                    {
                        if (!rowMap1.ContainsKey(kvp.Key))
                        {
                            Console.WriteLine($"Row non present in origin:");
                            PrintRowDetails(kvp.Value);
                        }
                    }

                    //CompareLabelsIRow(labels1, labels2);
                    Console.WriteLine("Extracting tabs and sections from forms...");
                    var labels1 = ExtractLabelsFromFormXml(formXml1);
                    var labels2 = ExtractLabelsFromFormXml(formXml2);
                    CompareLabels(labels1, labels2);

                    var trans1 = ExtractFieldTranslations(formXml1);
                    var trans2 = ExtractFieldTranslations(formXml2);
                    ExcelExport.ExportDifferences($"{entity}_{form}_diff.xlsx", fields1, fields2, trans1, trans2);
                    Console.WriteLine("\nComparison completed successfully.\n");

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{form}' for entity '{entity}': {ex.Message}");
                    Console.WriteLine($"Stack trace: {ex.StackTrace}");
                }
            }
        }
    }
}
