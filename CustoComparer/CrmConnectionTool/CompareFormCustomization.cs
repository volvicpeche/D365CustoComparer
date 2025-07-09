using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using static System.Net.WebRequestMethods;

namespace CrmConnectionTool
{
    public class CompareFormCustomization
    {
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

        // Méthode pour obtenir l'étiquette d'un nœud XML
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
                //XmlNode labelNode = fieldNode.SelectSingleNode("./labels/label");
                //if (labelNode != null && labelNode.Attributes?["description"] != null)
                //{
                //    customLabel = labelNode.Attributes["description"].Value;
                //    hasCustomLabel = true;
                //}
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

        // Méthode spécifique pour trouver les étiquettes personnalisées dans le XML du formulaire
        static Dictionary<string, string> ExtractCustomLabels(string formXml)
        {
            var customLabels = new Dictionary<string, string>();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(formXml);

            // Recherche des attributs avec des étiquettes personnalisées dans tous les contrôles
            XmlNodeList controlNodes = doc.SelectNodes("//control[@datafieldname]");
            foreach (XmlNode controlNode in controlNodes)
            {
                string fieldName = controlNode.Attributes?["datafieldname"]?.Value;
                if (string.IsNullOrEmpty(fieldName))
                    continue;

                // Recherche par attribut showlabel
                if (controlNode.Attributes?["showlabel"] != null)
                {
                    // Si showlabel est défini, il peut y avoir une étiquette personnalisée
                    XmlNode labelNode = controlNode.SelectSingleNode(".//labels/label[@description]");
                    if (labelNode != null && labelNode.Attributes["description"] != null)
                    {
                        customLabels[fieldName] = labelNode.Attributes["description"].Value;
                    }
                }

                // Recherche d'un nœud 'label' direct
                XmlNode directLabelNode = controlNode.SelectSingleNode("./label");
                if (directLabelNode != null && directLabelNode.Attributes?["description"] != null)
                {
                    customLabels[fieldName] = directLabelNode.Attributes["description"].Value;
                }

                // Recherche dans les paramètres de cellule
                XmlNodeList cellParams = controlNode.SelectNodes(".//cellLabelAttributes");
                foreach (XmlNode cellParam in cellParams)
                {
                    if (cellParam.InnerText?.Contains("customlabel") == true)
                    {
                        // Extraire la valeur personnalisée
                        try
                        {
                            XmlDocument paramDoc = new XmlDocument();
                            paramDoc.LoadXml(cellParam.InnerText);
                            XmlNode labelTextNode = paramDoc.SelectSingleNode("//customlabel");
                            if (labelTextNode != null && !string.IsNullOrEmpty(labelTextNode.InnerText))
                            {
                                customLabels[fieldName] = labelTextNode.InnerText;
                            }
                        }
                        catch
                        {
                            // Ignorer les erreurs de parsing
                        }
                    }
                }
            }

            return customLabels;
        }
        static void CompareFields(List<FormElement> fields1, List<FormElement> fields2)
        {
            // Identifier les champs ajoutés (présents dans fields2 mais pas dans fields1)
            var addedFields = fields2.Where(f2 => !fields1.Any(f1 => f1.Id == f2.Id)).ToList();

            // Identifier les champs supprimés (présents dans fields1 mais pas dans fields2)
            var removedFields = fields1.Where(f1 => !fields2.Any(f2 => f2.Id == f1.Id)).ToList();

            Console.WriteLine("\n--- Added Fields ---");
            foreach (var field in addedFields)
            {
                Console.WriteLine($"Field added: {field.DisplayName} in Tab: {field.TabName}, Section: {field.SectionName}");
            }

            Console.WriteLine("\n--- Removed Fields ---");
            foreach (var field in removedFields)
            {
                Console.WriteLine($"Field removed: {field.DisplayName} from Tab: {field.TabName}, Section: {field.SectionName}");
            }
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

        // Méthode pour comparer les étiquettes personnalisées des champs entre deux formulaires
        static void CompareFieldLabels(List<FormElement> fields1, List<FormElement> fields2)
        {
            Console.WriteLine("\n--- Changed Field Labels ---");

            // Comparer les étiquettes personnalisées
            foreach (var field1 in fields1)
            {
                var field2 = fields2.FirstOrDefault(f => f.FieldName == field1.FieldName);
                if (field2 != null)
                {
                    // Cas 1: Les deux champs ont des étiquettes personnalisées et elles sont différentes
                    if (field1.HasCustomLabel && field2.HasCustomLabel && field1.CustomLabel != field2.CustomLabel)
                    {
                        Console.WriteLine($"Field label changed for '{field1.FieldName}': '{field1.CustomLabel}' -> '{field2.CustomLabel}'");
                    }
                    // Cas 2: Le champ 1 a une étiquette personnalisée mais pas le champ 2
                    else if (field1.HasCustomLabel && !field2.HasCustomLabel)
                    {
                        Console.WriteLine($"Custom label removed for field '{field1.FieldName}' (was: '{field1.CustomLabel}')");
                    }
                    // Cas 3: Le champ 2 a une étiquette personnalisée mais pas le champ 1
                    else if (!field1.HasCustomLabel && field2.HasCustomLabel)
                    {
                        Console.WriteLine($"Custom label added for field '{field2.FieldName}': '{field2.CustomLabel}'");
                    }
                }
            }
        }

        static void CompareLabelsInXml(XmlDocument doc1, XmlDocument doc2)
        {
            // Cette méthode pourrait être étendue pour une comparaison plus spécifique des labels
            Console.WriteLine("\n--- Changes in label's description ---");

            XmlNodeList labelNodes1 = doc1.SelectNodes("//label[@description]");
            XmlNodeList labelNodes2 = doc2.SelectNodes("//label[@description]");

            if (labelNodes1 == null || labelNodes2 == null)
                return;

            // Créer des dictionnaires pour comparer facilement les labels par leur contexte
            Dictionary<string, string> labelDescs1 = new Dictionary<string, string>();
            Dictionary<string, string> labelDescs2 = new Dictionary<string, string>();

            foreach (XmlNode label in labelNodes1)
            {
                // Utiliser le chemin XPath comme "contexte" unique
                string context = GetLabelContext(label);
                string desc = label.Attributes["description"]?.Value;
                if (desc != null && !string.IsNullOrEmpty(context))
                    labelDescs1[context] = desc;
            }

            foreach (XmlNode label in labelNodes2)
            {
                string context = GetLabelContext(label);
                string desc = label.Attributes["description"]?.Value;
                if (desc != null && !string.IsNullOrEmpty(context))
                    labelDescs2[context] = desc;
            }

            // Comparer les descriptions de labels dans les mêmes contextes
            HashSet<string> allContexts = new HashSet<string>();
            foreach (var ctx in labelDescs1.Keys) allContexts.Add(ctx);
            foreach (var ctx in labelDescs2.Keys) allContexts.Add(ctx);

            foreach (string context in allContexts)
            {
                labelDescs1.TryGetValue(context, out string desc1);
                labelDescs2.TryGetValue(context, out string desc2);

                if (desc1 != desc2)
                {
                    Console.WriteLine($"Label changed for {context}:");
                    if (desc1 == null)
                        Console.WriteLine($"  - New label: '{desc2}'");
                    else if (desc2 == null)
                        Console.WriteLine($"  - Label remove (was: '{desc1}')");
                    else
                        Console.WriteLine($"  - Update: '{desc1}' -> '{desc2}'");
                }
            }
        }

        static string GetLabelContext(XmlNode labelNode)
        {
            // Trouver un identifiant de contexte pour le label (parent control/cell/etc avec ID)
            XmlNode parent = labelNode.ParentNode;
            while (parent != null)
            {
                if ((parent.Name == "control" || parent.Name == "cell" || parent.Name == "tab" || parent.Name == "section")
                    && parent.Attributes["id"] != null)
                {
                    return $"{parent.Name} {parent.Attributes["id"].Value}";
                }
                parent = parent.ParentNode;
            }
            return string.Empty;
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

        static string GetRowKey(XmlNode row)
        {
            // Extraire les informations pertinentes pour créer une clé unique et lisible
            XmlNode control = row.SelectSingleNode(".//control");
            XmlNode cell = row.SelectSingleNode(".//cell");
            XmlNode tab = row.SelectSingleNode("ancestor::tab");
            XmlNode section = row.SelectSingleNode("ancestor::section");

            if (control == null) return null;

            string fieldName = control.Attributes["datafieldname"]?.Value ?? "Unknown";
            string tabName = tab?.Attributes["name"]?.Value ?? "Unknown Tab";
            string sectionName = section?.Attributes["name"]?.Value ?? "Unknown Section";

            return $"{tabName} | {sectionName} | {fieldName}";
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

            Console.WriteLine($"  Tab: {tabName}");
            Console.WriteLine($"  Section: {sectionName}");
            Console.WriteLine($"  Field: {fieldName}");
            Console.WriteLine($"  ID: {cellId}");

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

                        Console.WriteLine($"    - Language: {langCode}");
                        Console.WriteLine($"      Description: {description}");
                        if (!string.IsNullOrEmpty(labelText))
                        {
                            Console.WriteLine($"      Text: {labelText}");
                        }
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

        static void CompareAttributes(XmlNode node1, XmlNode node2, string nodeType)
        {
            if (node1.Attributes == null || node2.Attributes == null)
                return;

            // Identifier le nœud (par son ID ou autre attribut reconnaissable)
            string id = node1.Attributes["id"]?.Value ?? "Inconnu";
            string name = node1.Attributes["datafieldname"]?.Value ?? id;

            // Comparer les attributs présents dans les deux nœuds
            HashSet<string> allAttributes = new HashSet<string>();
            foreach (XmlAttribute attr in node1.Attributes)
                allAttributes.Add(attr.Name);
            foreach (XmlAttribute attr in node2.Attributes)
                allAttributes.Add(attr.Name);

            bool differenceFound = false;
            StringBuilder differences = new StringBuilder();

            foreach (string attrName in allAttributes)
            {
                string value1 = node1.Attributes[attrName]?.Value;
                string value2 = node2.Attributes[attrName]?.Value;

                if (value1 != value2)
                {
                    if (!differenceFound)
                    {
                        differences.AppendLine($"Changes in {nodeType} '{name}' (id: {id}):");
                        differenceFound = true;
                    }

                    if (value1 == null)
                        differences.AppendLine($"  - Attribut '{attrName}' added: '{value2}'");
                    else if (value2 == null)
                        differences.AppendLine($"  - Attribut '{attrName}' removed (value was: '{value1}')");
                    else
                        differences.AppendLine($"  - Attribut '{attrName}' updated: '{value1}' -> '{value2}'");
                }
            }

            if (differenceFound)
                Console.WriteLine(differences.ToString());

            // Comparer récursivement les enfants qui contiennent des informations importantes
            CompareChildLabels(node1, node2, id);
        }

        static void CompareChildLabels(XmlNode node1, XmlNode node2, string parentId)
        {
            // Comparer les nœuds labels s'ils existent
            XmlNode labels1 = node1.SelectSingleNode(".//labels");
            XmlNode labels2 = node2.SelectSingleNode(".//labels");

            if (labels1 != null && labels2 != null)
            {
                XmlNodeList labelNodes1 = labels1.SelectNodes("./label");
                XmlNodeList labelNodes2 = labels2.SelectNodes("./label");

                Dictionary<string, string> labelMap1 = new Dictionary<string, string>();
                Dictionary<string, string> labelMap2 = new Dictionary<string, string>();

                // Collecter les descriptions par code de langue
                if (labelNodes1 != null)
                {
                    foreach (XmlNode label in labelNodes1)
                    {
                        string langCode = label.Attributes?["languagecode"]?.Value ?? "default";
                        string desc = label.Attributes?["description"]?.Value;
                        if (desc != null)
                            labelMap1[langCode] = desc;
                    }
                }

                if (labelNodes2 != null)
                {
                    foreach (XmlNode label in labelNodes2)
                    {
                        string langCode = label.Attributes?["languagecode"]?.Value ?? "default";
                        string desc = label.Attributes?["description"]?.Value;
                        if (desc != null)
                            labelMap2[langCode] = desc;
                    }
                }

                // Comparer les descriptions
                HashSet<string> allLangCodes = new HashSet<string>();
                foreach (var code in labelMap1.Keys) allLangCodes.Add(code);
                foreach (var code in labelMap2.Keys) allLangCodes.Add(code);

                bool differenceFound = false;
                StringBuilder differences = new StringBuilder();

                foreach (string code in allLangCodes)
                {
                    labelMap1.TryGetValue(code, out string desc1);
                    labelMap2.TryGetValue(code, out string desc2);

                    if (desc1 != desc2)
                    {
                        if (!differenceFound)
                        {
                            differences.AppendLine($"label changed for element ID '{parentId}':");
                            differenceFound = true;
                        }

                        if (desc1 == null)
                            differences.AppendLine($"  - Label added for language '{code}': '{desc2}'");
                        else if (desc2 == null)
                            differences.AppendLine($"  - Label deleted for language '{code}' (valeur était: '{desc1}')");
                        else
                            differences.AppendLine($"  - Label updated for language '{code}': '{desc1}' -> '{desc2}'");
                    }
                }

                if (differenceFound)
                    Console.WriteLine(differences.ToString());
            }
        }

        static void CompareNodesByAttribute(XmlDocument doc1, XmlDocument doc2, string xpathQuery, string idAttribute)
        {
            XmlNodeList nodes1 = doc1.SelectNodes(xpathQuery);
            XmlNodeList nodes2 = doc2.SelectNodes(xpathQuery);

            if (nodes1 == null || nodes2 == null)
                return;

            Dictionary<string, XmlNode> nodeMap1 = new Dictionary<string, XmlNode>();
            Dictionary<string, XmlNode> nodeMap2 = new Dictionary<string, XmlNode>();

            // Indexer les nœuds par ID
            foreach (XmlNode node in nodes1)
            {
                string id = node.Attributes?[idAttribute]?.Value;
                if (!string.IsNullOrEmpty(id))
                    nodeMap1[id] = node;
            }

            foreach (XmlNode node in nodes2)
            {
                string id = node.Attributes?[idAttribute]?.Value;
                if (!string.IsNullOrEmpty(id))
                    nodeMap2[id] = node;
            }

            // Comparer les nœuds avec le même ID
            foreach (var kvp in nodeMap1)
            {
                string id = kvp.Key;
                XmlNode node1 = kvp.Value;

                if (nodeMap2.TryGetValue(id, out XmlNode node2))
                {
                    CompareAttributes(node1, node2, xpathQuery.Replace("//", ""));
                }
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

        public static void CompareForm(IOrganizationService service1, IOrganizationService service2, List<(string EntityName, string FormName)> formsToCompare )
        {
            foreach (var (entity, form) in formsToCompare)
            {
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
                        Console.WriteLine("The forms are identical. No changes detected.");
                        continue;
                    }

                    Console.WriteLine("Extracting fields from first form...");
                    var fields1 = ExtractFieldsFromFormXml(formXml1);
                    var rows1 = ExtractRowsFromFormXml(formXml1);
                    Console.WriteLine($"Found {fields1.Count} fields in first form.");

                    Console.WriteLine("Extracting fields from second form...");
                    var fields2 = ExtractFieldsFromFormXml(formXml2);
                    var rows2 = ExtractRowsFromFormXml(formXml2);
                    Console.WriteLine($"Found {fields2.Count} fields in second form.");

                    //review 1
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

                    Console.WriteLine("Extracting tabs and sections from forms...");
                    var labels1 = ExtractLabelsFromFormXml(formXml1);
                    var labels2 = ExtractLabelsFromFormXml(formXml2);
                    CompareLabels(labels1, labels2);

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