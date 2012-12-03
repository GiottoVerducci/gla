using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace GLA
{
    public static class ExcelLoader
    {
        public static Dictionary<Peuple, Army> LoadAll(string path)
        {
            var result = new Dictionary<Peuple, Army>();

            LoadFile(path + "figurine_taille.xlsx", "figurine_taille", Army.AllArmies, Army.Sizes);
            LoadFile(path + "figurine_peuple.xlsx", "figurine_peuple", Army.AllArmies, Army.Peuples);
            LoadFile(path + "figurine_famille.xlsx", "figurine_famille", Army.AllArmies, Army.Families);
            LoadFile(path + "figurine_rang.xlsx", "figurine_rang", Army.AllArmies, Army.Ranks);
            LoadFile(path + "figurine_capacite.xlsx", "figurine_capacite", Army.AllArmies, Army.AllArmies.Capacities);
            LoadFile(path + "artefact.xlsx", "artefact", Army.AllArmies, Army.AllArmies.Artifacts);
            LoadFile(path + "magie.xlsx", "magie", Army.AllArmies, Army.AllArmies.Spells);
            LoadFile(path + "magie_voie.xlsx", "magie_voie", Army.AllArmies, Army.MagicPaths);
            LoadFile(path + "miracle.xlsx", "miracle", Army.AllArmies, Army.AllArmies.Miracles);
            LoadFile(path + "figurine.xlsx", "figurine", Army.AllArmies, Army.AllArmies.Units);

            LoadFile(path + "figurine_lienartefact.xlsx", "figurine_lienartefact", Army.AllArmies, Army.AllArmies.UnitArtifactLinks);
            LoadFile(path + "figurine_liencapacite.xlsx", "figurine_liencapacite", Army.AllArmies, Army.AllArmies.UnitCapacityLinks);
            LoadFile(path + "figurine_lienfamille.xlsx", "figurine_lienfamille", Army.AllArmies, Army.AllArmies.UnitFamilyLinks);
            LoadFile(path + "figurine_lienmiracle.xlsx", "figurine_lienmiracle", Army.AllArmies, Army.AllArmies.UnitMiracleLinks);
            LoadFile(path + "figurine_liensort.xlsx", "figurine_liensort", Army.AllArmies, Army.AllArmies.UnitSpellLinks);
            LoadFile(path + "magie_voie_lienpeuple.xlsx", "magie_voie_lienpeuple", Army.AllArmies, Army.AllArmies.PeupleMagicPathLinks);

            foreach (var unit in Army.AllArmies.Units.Values)
                unit.Resolve();
            foreach (var artifact in Army.AllArmies.Artifacts.Values)
                artifact.Resolve();
            foreach (var spell in Army.AllArmies.Spells.Values)
                spell.Resolve();
            foreach (var miracle in Army.AllArmies.Miracles.Values)
                miracle.Resolve();
            foreach (var capacity in Army.AllArmies.Capacities.Values)
                capacity.Resolve();
            foreach(var unitArtifactLink in Army.AllArmies.UnitArtifactLinks)
                unitArtifactLink.Resolve();
            foreach (var unitCapacityLink in Army.AllArmies.UnitCapacityLinks)
                unitCapacityLink.Resolve();
            foreach (var unitFamilyLink in Army.AllArmies.UnitFamilyLinks)
                unitFamilyLink.Resolve();
            foreach (var unitMiracleLink in Army.AllArmies.UnitMiracleLinks)
                unitMiracleLink.Resolve();
            foreach (var unitSpellLink in Army.AllArmies.UnitSpellLinks)
                unitSpellLink.Resolve();
            foreach (var peupleMagicPathLink in Army.AllArmies.PeupleMagicPathLinks)
                peupleMagicPathLink.Resolve();

            foreach (var unit in Army.AllArmies.Units.Values)
            {
                Army army;
                if (!result.TryGetValue(unit.Peuple.Entity, out army))
                {
                    army = new Army();
                    result.Add(unit.Peuple.Entity, army);
                }

                army.Units.Add(unit.Id, unit);
                foreach (var artifact in unit.AssociatedArtifacts)
                    army.Artifacts[artifact.Id] = artifact;
                foreach (var spell in unit.AssociatedSpells)
                    army.Spells[spell.Id] = spell;
                foreach (var miracle in unit.AssociatedMiracles)
                    army.Miracles[miracle.Id] = miracle;
                foreach (var capacity in unit.AssociatedCapacities)
                    army.Capacities[capacity.Id] = capacity;
            }

            var allUnits = Army.AllArmies.Units.Values.OrderBy(u => u.FullName).ToList();
            var groupedUnits = Unit.GetGroupedUnits(allUnits);
            foreach (var groupedUnit in groupedUnits)
            {
                foreach(var unit in groupedUnit.Value)
                {
                    unit.ReferenceUnit = groupedUnit.Value[0];
                    unit.VariationCount = groupedUnit.Value.Count;
                }
            }

            return result;
        }

        private static void LoadFile<T>(string filename, string sheetName, Army army, Dictionary<int, T> dictionary)
            where T : ArmyPart, new()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                var sheets = workbookPart.Workbook.Descendants<Sheet>();

                foreach (var sheet in sheets)
                {
                    var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    var worksheet = worksheetPart.Worksheet;

                    var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
                    if (sheet.Name.ToString() == sheetName)
                    {
                        foreach (var item in SheetDataLoader<T>.Read(workbookPart, sheetData, army))
                            dictionary.Add(item.Id, item);
                        Console.WriteLine("{0}: {1} items loaded.", typeof(T), dictionary.Count);
                    }
                }
            }
        }

        private static void LoadFile<T>(string filename, string sheetName, Army army, List<T> list)
            where T : ArmyPart, new()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                var sheets = workbookPart.Workbook.Descendants<Sheet>();

                foreach (var sheet in sheets)
                {
                    var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    var worksheet = worksheetPart.Worksheet;

                    var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
                    if (sheet.Name.ToString() == sheetName)
                    {
                        list.AddRange(SheetDataLoader<T>.Read(workbookPart, sheetData, army));
                        Console.WriteLine("{0}: {1} items loaded.", typeof(T), list.Count);
                    }
                }
            }
        }

        public static Army ReadArmy(string filename)
        {
            var result = new Army();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                // Console.WriteLine("Worksheet Parts: " + workbookPart.WorksheetParts.Count());

                var sheets = workbookPart.Workbook.Descendants<Sheet>();

                var i = 0;
                foreach (var sheet in sheets)
                {
                    var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    var worksheet = worksheetPart.Worksheet;

                    // Console.WriteLine("\tWorksheet Part " + i + " : " + sheet.Name);

                    var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
                    switch (sheet.Name.ToString())
                    {
                        case "Général":
                            ReadGeneralProperties(workbookPart, sheetData, result.GeneralProperties);
                            Console.WriteLine("{0} general properties loaded.", result.GeneralProperties.Count);
                            break;
                        //case "Unités":
                        //    result.Units.AddRange(SheetDataLoader<Unit>.Read(workbookPart, sheetData, result));
                        //    Console.WriteLine("{0} units loaded.", result.Units.Count);
                        //    break;
                        //case "Spécial":
                        //    result.Specials.AddRange(SheetDataLoader<Special>.Read(workbookPart, sheetData, result));
                        //    Console.WriteLine("{0} specials loaded.", result.Specials.Count);
                        //    break;
                        //case "Artefacts":
                        //    result.Artefacts.AddRange(SheetDataLoader<Artefact>.Read(workbookPart, sheetData, result));
                        //    Console.WriteLine("{0} artefacts loaded.", result.Artefacts.Count);
                        //    break;
                        //case "Sortilèges":
                        //    result.Sortileges.AddRange(SheetDataLoader<Sortilege>.Read(workbookPart, sheetData, result));
                        //    Console.WriteLine("{0} artefacts loaded.", result.Sortileges.Count);
                        //    break;
                        //case "Factions":
                        //    result.FactionElements.AddRange(SheetDataLoader<FactionElement>.Read(workbookPart, sheetData, result));
                        //    Console.WriteLine("{0} factions elements loaded.", result.FactionElements.Count);
                        //    break;
                        //case "Factions - combattants":
                        //    result.FactionCombattantElements.AddRange(SheetDataLoader<FactionCombattantElement>.Read(workbookPart, sheetData, result));
                        //    Console.WriteLine("{0} factions combattants elements loaded.", result.FactionCombattantElements.Count);
                        //    break;
                    }

                    //Console.WriteLine("\tWorksheet - Sheets: " + worksheet.Elements<SheetData>().Count());

                    //var j = 0;
                    //foreach (var sheetData in worksheet.Elements<SheetData>())
                    //{
                    //    Console.WriteLine("\t\tWorksheet - Sheet " + j);

                    //    foreach (Row r in sheetData.Elements<Row>())
                    //    {
                    //        foreach (Cell c in r.Elements<Cell>())
                    //        {
                    //            if (c.CellValue != null)
                    //            {
                    //                string text = GetCellValue(c, workbookPart);
                    //                Console.Write(text + "\t");
                    //            }
                    //        }
                    //        Console.WriteLine();
                    //    }
                    //    ++j;
                    //}
                    ++i;
                }
            }
            result.Resolve();
            return result;
        }

        private static void ReadGeneralProperties(WorkbookPart workbookPart, SheetData sheetData, Dictionary<string, string> generalProperties)
        {
            foreach (Row row in sheetData.Elements<Row>())
            {
                bool firstCell = true;
                string propertyName = null;
                foreach (Cell c in row.Elements<Cell>())
                {
                    if (c.CellValue != null)
                    {
                        if (firstCell)
                        {
                            propertyName = GetCleanText(GetCellValue(c, workbookPart));
                            firstCell = false;
                        }
                        else
                        {
                            generalProperties.Add(propertyName, GetCleanText(GetCellValue(c, workbookPart)));
                            break;
                        }
                    }
                }
            }
        }

        public static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            string value = null;
            // If the cell does not exist, return an empty string:
            if (cell != null)
            {
                value = cell.InnerText;

                // If the cell represents a numeric value, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and Booleans
                // individually. For shared strings, the code looks up the 
                // corresponding value in the shared string table. For Booleans, 
                // the code converts the value into the words TRUE or FALSE.
                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the shared 
                            // strings table.
                            var stringTable = workbookPart.
                                GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            // If the shared string table is missing, something is 
                            // wrong. Return the index that you found in the cell.
                            // Otherwise, look up the correct text in the table.
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.
                                    ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }

        public static string GetCleanText(string text)
        {
            text = text.Trim();
            text = text.Replace("_x000D_", null);
            while (text.Length > 0 && text[text.Length - 1] == '\n')
            {
                text = text.Substring(0, text.Length - 1);
                text = text.TrimEnd();
            }

            text = text
                .Replace("« ", "«" + Const.NBSP)
                .Replace(" »", Const.NBSP + "»")
                .Replace("&nbsp;", Const.NBSP.ToString());
            return text;
        }

    }

}