using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CrmConnectionTool
{
    public class CompareViewCustomization
    {
        static void CompareXmlNodes(string formXml1, string formXml2)
        {
            XmlDocument doc1 = new XmlDocument();
            XmlDocument doc2 = new XmlDocument();

            try
            {
                doc1.LoadXml(formXml1);
                doc2.LoadXml(formXml2);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors du chargement XML: {ex.Message}");
                return;
            }

            Console.WriteLine("\n--- Comparaison détaillée des rows ---");

            // Extraire toutes les rows des deux documents
            XmlNodeList rows1 = doc1.SelectNodes("//row");
            XmlNodeList rows2 = doc2.SelectNodes("//row");

            if (rows1 == null || rows2 == null)
            {
                Console.WriteLine("Aucun élément 'row' trouvé dans l'un des documents XML.");
                return;
            }

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

            // Comparer les rows avec le même ID de cellule
            Console.WriteLine("\nComparaison des rows:");
            foreach (var kvp in rowMap1)
            {
                string cellId = kvp.Key;
                XmlNode row1 = kvp.Value;

                if (rowMap2.TryGetValue(cellId, out XmlNode row2))
                {
                    // Comparer les attributs de la row et de ses composants
                    CompareRowStructures(row1, row2, cellId);
                }
                else
                {
                    Console.WriteLine($"Row avec l'ID de cellule '{cellId}' présente uniquement dans le premier formulaire.");
                }
            }

            // Identifier les rows présentes uniquement dans le deuxième formulaire
            foreach (var kvp in rowMap2)
            {
                if (!rowMap1.ContainsKey(kvp.Key))
                {
                    Console.WriteLine($"Row avec l'ID de cellule '{kvp.Key}' présente uniquement dans le deuxième formulaire.");
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
                        differences.AppendLine($"  - Attribut '{attrName}' ajouté dans le 2ème formulaire: '{value2}'");
                    else if (value2 == null)
                        differences.AppendLine($"  - Attribut '{attrName}' supprimé dans le 2ème formulaire (valeur était: '{value1}')");
                    else
                        differences.AppendLine($"  - Attribut '{attrName}' modifié: '{value1}' -> '{value2}'");
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
    }
}
