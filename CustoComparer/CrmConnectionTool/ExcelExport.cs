using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace CrmConnectionTool
{
    public static class ExcelExport
    {
        static Cell CreateCell(string text)
        {
            return new Cell { DataType = CellValues.String, CellValue = new CellValue(text ?? string.Empty) };
        }

        static void FillSheet(SheetData sheetData, IEnumerable<List<string>> rows)
        {
            foreach (var rowValues in rows)
            {
                var row = new Row();
                foreach (var value in rowValues)
                {
                    row.Append(CreateCell(value));
                }
                sheetData.Append(row);
            }
        }

        public static void ExportDifferences(string path,
            List<Program.FormElement> fields1,
            List<Program.FormElement> fields2,
            Dictionary<string, Dictionary<string, string>> trans1,
            Dictionary<string, Dictionary<string, string>> trans2)
        {
            using (var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new Workbook();
                var sheets = wbPart.Workbook.AppendChild(new Sheets());

                uint sheetId = 1;
                SheetData AddSheet(string name)
                {
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    var sd = new SheetData();
                    wsPart.Worksheet = new Worksheet(sd);
                    sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = name });
                    return sd;
                }

                var addedRemoved = AddSheet("Fields");
                var labelChanges = AddSheet("Labels");
                var translationChanges = AddSheet("Translations");
                var movedFields = AddSheet("TabSection");

                // compute differences
                var added = new List<Program.FormElement>();
                var removed = new List<Program.FormElement>();
                foreach (var f2 in fields2)
                {
                    if (!fields1.Exists(f1 => f1.FieldName == f2.FieldName))
                        added.Add(f2);
                }
                foreach (var f1 in fields1)
                {
                    if (!fields2.Exists(f2 => f2.FieldName == f1.FieldName))
                        removed.Add(f1);
                }

                var addedRemovedRows = new List<List<string>>();
                addedRemovedRows.Add(new List<string>{"Field","Change","Tab","Section"});
                foreach (var f in added)
                    addedRemovedRows.Add(new List<string>{f.FieldName,"Added",f.TabName,f.SectionName});
                foreach (var f in removed)
                    addedRemovedRows.Add(new List<string>{f.FieldName,"Removed",f.TabName,f.SectionName});
                FillSheet(addedRemoved, addedRemovedRows);

                var labelRows = new List<List<string>> { new List<string>{"Field","Old Label","New Label"} };
                foreach (var f1 in fields1)
                {
                    var f2 = fields2.Find(x => x.FieldName == f1.FieldName);
                    if (f2 != null && f1.CustomLabel != f2.CustomLabel)
                    {
                        labelRows.Add(new List<string>{f1.FieldName,f1.CustomLabel ?? string.Empty,f2.CustomLabel ?? string.Empty});
                    }
                }
                FillSheet(labelChanges, labelRows);

                var translationRows = new List<List<string>> { new List<string>{"Field","Language","Old","New"} };
                foreach (var kvp in trans1)
                {
                    if (!trans2.TryGetValue(kvp.Key, out var t2)) continue;
                    var langs = new HashSet<string>(kvp.Value.Keys);
                    foreach (var l in t2.Keys) langs.Add(l);
                    foreach (var lang in langs)
                    {
                        kvp.Value.TryGetValue(lang, out var v1);
                        t2.TryGetValue(lang, out var v2);
                        if (v1 != v2)
                            translationRows.Add(new List<string>{kvp.Key, lang, v1 ?? string.Empty, v2 ?? string.Empty});
                    }
                }
                FillSheet(translationChanges, translationRows);

                var moveRows = new List<List<string>> { new List<string>{"Field","Old Tab","Old Section","New Tab","New Section"} };
                foreach (var f1 in fields1)
                {
                    var f2 = fields2.Find(x => x.FieldName == f1.FieldName);
                    if (f2 != null && (f1.TabName != f2.TabName || f1.SectionName != f2.SectionName))
                    {
                        moveRows.Add(new List<string>{f1.FieldName,f1.TabName,f1.SectionName,f2.TabName,f2.SectionName});
                    }
                }
                FillSheet(movedFields, moveRows);

                wbPart.Workbook.Save();
            }
        }
    }
}

