using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;

namespace GLA
{
    public static class Pdf
    {
        private static Font _titleFont;
        private static Font _sectionTitleFont;
        private static Font _subsectionTitleFont;
        private static Font _nameFont;
        private static Font _textFont;
        private static Font _whiteTextFont;
        private static Font _boldTextFont;
        private static Font _whiteBoldTextFont;
        private static Font _italicTextFont;
        private static Font _specialFont;
        private static Font _reservedToFont;

        private static BaseFont _allTextFont;
        private static BaseFont _allBoldTextFont;
        private static float _sectionTitleIndent;
        private const float _cellSpace = 20f;

        //http://www.mikesdotnetting.com/Article/80/Create-PDFs-in-ASP.NET-getting-started-with-iTextSharp
        public static void Write(Army army, Peuple peuple, string fileName, string footerText, string headerLeftText, string headerRightText, string watermarkImagePath, int watermarkOpacity)
        {
            //BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
            //Font times = new Font(bfTimes, 12, Font.ITALIC, BaseColor.RED);

            //var arialBoldBase = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, true);
            //var sectionTitleFont = new Font(arialBoldBase, 16, Font.NORMAL, new BaseColor(128, 0, 128)); //FontFactory.GetFont("dax-black");
            //var boldArialText = new Font(arialBoldBase, 10, Font.NORMAL);

            //_sectionTitleFont = FontFactory.GetFont("Arial", 16, Font.BOLD, new BaseColor(128, 0, 128));
            //_subsectionTitleFont = FontFactory.GetFont("Arial", 12, Font.BOLD);
            //_nameFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            //_textFont = FontFactory.GetFont("Arial", 10);
            //_specialFont = FontFactory.GetFont("Arial", 10, Font.BOLDITALIC);
            //_reservedToFont = FontFactory.GetFont("Arial", 10, Font.BOLD, new BaseColor(128, 0, 128));
            BaseFont antiqueFont = BaseFont.CreateFont(@"Pdf\casablanca-antique-plain.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont antiqueBoldFont = BaseFont.CreateFont(@"Pdf\CasablancaAntique-Bold.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            _sectionTitleFont = new Font(antiqueFont, 16, Font.NORMAL, new BaseColor(128, 0, 128));
            _subsectionTitleFont = new Font(antiqueBoldFont, 12, Font.NORMAL);
            _allTextFont = BaseFont.CreateFont(@"Pdf\itc-goudy-sans-lt-medium.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont allItalicTextFont = BaseFont.CreateFont(@"Pdf\itc-goudy-sans-lt-medium-italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            _allBoldTextFont = BaseFont.CreateFont(@"Pdf\itc-goudy-sans-lt-bold.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont allBoldItalicTextFont = BaseFont.CreateFont(@"Pdf\itc-goudy-sans-lt-bold-italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            _titleFont = new Font(antiqueFont, 60);
            _nameFont = new Font(_allBoldTextFont, 10);
            _textFont = new Font(_allTextFont, 10);
            _whiteTextFont = new Font(_allTextFont, 10, Font.NORMAL, BaseColor.WHITE);
            _italicTextFont = new Font(allItalicTextFont, 10);
            _boldTextFont = new Font(_allBoldTextFont, 10);
            _whiteBoldTextFont = new Font(_allBoldTextFont, 10, Font.NORMAL, BaseColor.WHITE);
            _specialFont = new Font(allBoldItalicTextFont, 10);
            _reservedToFont = new Font(_allBoldTextFont, 10, Font.NORMAL, new BaseColor(128, 0, 128));
            _sectionTitleIndent = 50f;

            Image watermarkImage = string.IsNullOrWhiteSpace(watermarkImagePath)
                ? null
                : Image.GetInstance(watermarkImagePath);

            var doc = new Document();
            //use a variable to let my code fit across the page...
            var writer = PdfWriter.GetInstance(doc, new FileStream(fileName, FileMode.Create));
            writer.PageEvent = new PdfFooterHelper(footerText, headerLeftText, headerRightText, watermarkImage, watermarkOpacity);
            doc.Open();

            var columnManager = new ColumnManager(doc, writer);

            #region Title page
            string title;
            //if (army.GeneralProperties.TryGetValue("Titre", out title))
            {
                title = "\n\nLivre d'armée\n" + peuple._name;
                var paragraph = new Paragraph(title, _titleFont) { Alignment = Element.ALIGN_CENTER };
                //columnManager.AddUnbreakable(paragraph);
                //columnManager.Render(1);
                doc.Add(paragraph);
                columnManager.EndPage();
            }

            //doc.Add(new Paragraph("My first PDF"));

            //writer.SetLinearPageMode();
            //var toc = new List<TableOfContentsEntry>
            //{
            //    new TableOfContentsEntry("Execution Page", writer.CurrentPageNumber.ToString())
            //};

            //Chapter executionPage = new Chapter(new Paragraph("Execution Page", FontFactory.GetFont("Arial", 18, Font.BOLD)), 3);


            //PdfPTable table = new PdfPTable(2);

            //PdfPCell titleCell = new PdfPCell(table);
            //titleCell.Border = Rectangle.NO_BORDER;
            //titleCell.Padding = 4f;
            //titleCell.Phrase = new Phrase("Title");
            //table.AddCell(titleCell);

            //PdfPCell titleContentCell = new PdfPCell(table);
            //titleContentCell.Border = Rectangle.NO_BORDER;
            //titleContentCell.Padding = 4f;
            //titleContentCell.Phrase = new Phrase("Search Engine");
            //table.AddCell(titleContentCell);

            //PdfPCell nameCell = new PdfPCell(table);
            //nameCell.Border = Rectangle.NO_BORDER;
            //nameCell.Padding = 4f;
            //nameCell.Phrase = new Phrase("Name");
            //table.AddCell(nameCell);

            //PdfPCell nameContentCell = new PdfPCell(table);
            //nameContentCell.Border = Rectangle.NO_BORDER;
            //nameContentCell.Padding = 4f;
            //nameContentCell.Phrase = new Phrase("Google");

            //table.AddCell(nameContentCell);

            //table.SetExtendLastRow(false, false);
            //table.ElementComplete = true;

            //executionPage.Add(table);

            //doc.Add(executionPage);


            //var table = new PdfPTable(3);
            //table.DefaultCell.BorderWidth = 0f;
            //var cell = new PdfPCell(new Phrase("Header spanning 3 columns"));
            //cell.Colspan = 3;
            //cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
            //table.AddCell(cell);
            //table.AddCell("Col 1 Row 1");
            //table.AddCell("Col 2 Row 1");
            //table.AddCell("Col 3 Row 1");
            //table.AddCell("Col 1 Row 2");
            //table.AddCell("Col 2 Row 2");
            //table.AddCell("Col 3 Row 2");
            //doc.Add(table);


            //Chunk chunk = new Chunk("Setting the Font", FontFactory.GetFont("dax-black"));
            //chunk.SetUnderline(0.5f, -1.5f);

            //Phrase phrase = new Phrase();
            //for (int i = 1; i < 4; i++)
            //{
            //    phrase.Add(chunk);
            //}

            //MultiColumnText columns = new MultiColumnText();
            ////float left, float right, float gutterwidth, int numcolumns
            //columns.AddRegularColumns(36f, doc.PageSize.Width - 36f, 24f, 2);

            //PdfContentByte cb = writer.DirectContent;
            //var ct = new ColumnText(cb);
            //ct.Alignment = Element.ALIGN_JUSTIFIED;
            #endregion

            #region Lists page
            Console.WriteLine("Lists");
            var units = army.Units.Values.Where(u => !u.IsPersonnage).ToList();
            if (units.Count > 0)
            {
                var paragraph = GetSectionTitleParagraph("Liste des coûts des combattants individuels\n");
                //ct.AddText(paragraph);
                //RenderColumns(doc, ct, 1);
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var unit in units.OrderBy(u => u.FullName))
                {
                    paragraph = new Paragraph();
                    //var unitName = new Anchor(unit.FullName, _nameFont) { Reference = "#" + unit.FullName };
                    var cost = new Chunk(string.Format(" : {0} PA\n", unit._points), _textFont);
                    //paragraph.Add(new Phrase { unitName, cost });
                    paragraph.Add(GetUnitNameReference(unit, _nameFont));
                    paragraph.Add(cost);
                    columnManager.AddUnbreakable(paragraph);
                }
                //ct.AddText(paragraph);
                //RenderColumns(doc, ct, 2);
                columnManager.Render(2);
                columnManager.EndPage();
            }

            var champions = army.Units.Values.Where(u => u.IsPersonnage).ToList();
            if (champions.Count > 0)
            {
                var paragraph = GetSectionTitleParagraph("Liste des coûts des Champions\n");
                //ct.AddText(paragraph);
                //RenderColumns(doc, ct, 1);
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var unit in champions.OrderBy(u => u.FullName))
                {
                    paragraph = new Paragraph();
                    //var unitName = new Anchor(unit.FullName, _nameFont) { Reference = "#" + unit.FullName };
                    var cost = new Chunk(string.Format(" : {0} PA\n", unit._points), _textFont);
                    //paragraph.Add(new Phrase { unitName, cost });
                    paragraph.Add(GetUnitNameReference(unit, _nameFont));
                    paragraph.Add(cost);
                    columnManager.AddUnbreakable(paragraph);
                }
                //ct.AddText(paragraph);
                //RenderColumns(doc, ct, 2);
                columnManager.Render(2);
                columnManager.EndPage();
                //EndPage();
            }
            #endregion

            #region Profils détaillés des combattants individuels

            var nonMagicalUnits = units.Where(u => !u.IsMagical && !u.IsDivine).ToList();
            if (nonMagicalUnits.Count > 0)
            {
                var paragraph = GetSectionTitleParagraph("Profils détaillés des combattants individuels\n");
                //ct.AddText(paragraph);
                //RenderColumns(doc, ct, 1);
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var unit in nonMagicalUnits.OrderBy(u => u.FullName))
                {
                    var paragraphs = GetUnitParagraphs(army, unit, columnManager.GetColumnWidth(2));
                    columnManager.AddUnbreakable(AddSeparatorParagraph(paragraphs));
                }
                //ct.AddText(paragraph);
                //RenderColumns(doc, ct, 2);
                columnManager.Render(2);
                columnManager.EndPage();
            }
            #endregion
            //            goto end;

            #region Profils détaillés des Magiciens
            var magicalUnits = units.Where(u => u.IsMagical).ToList();
            if (magicalUnits.Count > 0)
            {
                var paragraph = GetSectionTitleParagraph("Profils détaillés des Magiciens\n");
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var unit in magicalUnits.OrderBy(u => u.FullName))
                {
                    var paragraphs = GetUnitParagraphs(army, unit, columnManager.GetColumnWidth(2));
                    columnManager.AddUnbreakable(AddSeparatorParagraph(paragraphs));
                }
                columnManager.Render(2);
                columnManager.EndPage();
            }
            #endregion
            #region Profils détaillés des Fidèles
            var divineUnits = units.Where(u => u.IsDivine).ToList();
            if (divineUnits.Count > 0)
            {
                var paragraph = GetSectionTitleParagraph("Profils détaillés des Fidèles\n");
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var unit in divineUnits.OrderBy(u => u.FullName))
                {
                    var paragraphs = GetUnitParagraphs(army, unit, columnManager.GetColumnWidth(2));
                    columnManager.AddUnbreakable(paragraphs);
                }
                columnManager.Render(2);
                columnManager.EndPage();
            }
            #endregion

            #region Profils détaillés des Champions
            if (champions.Count > 0)
            {
                var paragraph = GetSectionTitleParagraph("Profils détaillés des Champions\n");
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var unit in champions.OrderBy(u => u.FullName))
                {
                    var paragraphs = GetUnitParagraphs(army, unit, columnManager.GetColumnWidth(2));
                    //if (unit.Id == 10058)
                    //    columnManager.AddSpUnbreakable(AddSeparatorParagraph(paragraphs));
                    //else
                    columnManager.AddUnbreakable(AddSeparatorParagraph(paragraphs));
                }
                columnManager.Render(2);
                columnManager.EndPage();
            }
            #endregion

            #region Liste des artefacts réservés
            var nonUniqueArtefacts = army.Artifacts.Values.Where(a => a.GroupedAssociatedUnits.Count > 1).ToList();
            if (champions.Count > 0)
            {
                var paragraph = GetSectionTitleParagraph("Liste des artefacts réservés\n");
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var artefact in nonUniqueArtefacts.OrderBy(u => u._name))
                {
                    var paragraphs = GetArtefactParagraph(artefact);
                    columnManager.AddUnbreakable(paragraphs);
                }
                columnManager.Render(2);
                columnManager.EndPage();
            }
            #endregion

            #region Sortilèges
            //var spellsByPath = army.Spells.Values.GroupBy(s => s.Voies.First());
            var magicPaths = peuple.AssociatedMagicPaths;
            //foreach (var pathGroup in spellsByPath)
            //{
            //    var voie = pathGroup.Key;
            foreach(var voie in magicPaths)
                {
                var paragraph = GetSectionTitleParagraphWithLink(new[] { 1 },
                    "Liste des sortilèges de la voie ",
                    voie.QualifiedName,
                    Environment.NewLine);
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);
                int i = 0;
                //foreach (var sortilege in pathGroup)
                var sortileges = army.Spells.Values.Where(s => s.Voies.Contains(voie._name)).ToList();
                foreach(var sortilege in sortileges)
                {
                    paragraph = GetSpellParagraph(sortilege, columnManager.GetColumnWidth(2));
                    if (++i < sortileges.Count)
                        paragraph.Add(new LineSeparator(1f, 100f, BaseColor.BLACK, Element.ALIGN_CENTER, -1));
                    columnManager.AddUnbreakable(paragraph);
                }
                columnManager.Render(2);
                columnManager.EndPage();
            }
            #endregion

            #region Miracles
            var miraclesByCult = army.Miracles.Values.GroupBy(s => s.Cults.First());
            foreach (var pathGroup in miraclesByCult)
            {
                var cult = pathGroup.Key;
                var paragraph = GetSectionTitleParagraphWithLink(new[] { 1 },
                    cult.StartsWithVowel() ? "Liste des miracles du culte de l'" : "Liste des miracles de la voie du ",
                    cult,
                    Environment.NewLine);
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);
                int i = 0;
                foreach (var miracle in pathGroup)
                {
                    paragraph = GetMiracleParagraph(miracle, columnManager.GetColumnWidth(2));
                    if (++i < pathGroup.Count())
                        paragraph.Add(new LineSeparator(1f, 100f, BaseColor.BLACK, Element.ALIGN_CENTER, -1));
                    columnManager.AddUnbreakable(paragraph);
                }
                columnManager.Render(2);
                columnManager.EndPage();
            }
            #endregion

            #region Factions
            foreach (var faction in army.Factions)
            {
                var paragraph = GetSectionTitleParagraph("Faction : {0}\n", faction.Nom);
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                paragraph = GetArmyPartParagraph();
                var coutAffiliation = GetCoutText(faction.Affiliation._PA, faction.Affiliation._PAChampion);
                paragraph.Add(new Chunk(JoinClean(" ", new[] { faction.Affiliation._Nom, "(affiliation)", coutAffiliation }) + "\n", _nameFont));
                paragraph.Add(new Chunk(faction.Affiliation._Effet + Environment.NewLine, _textFont));
                columnManager.AddUnbreakable(paragraph);

                foreach (var solo in faction.Solos)
                {
                    paragraph = GetArmyPartParagraph();
                    var coutSolo = GetCoutText(solo._PA, solo._PAChampion);
                    //var reserve = !string.IsNullOrWhiteSpace(solo._Reserve) ? solo._Reserve : "Tous";
                    //paragraph.Add(new Chunk(string.Format("{0} (solo) : {1} {2} : ", solo._nom, reserve, coutSolo), _nameFont));
                    paragraph.Add(new Chunk(string.Format("{0} (solo) / ", solo._Nom), _nameFont));
                    if (solo.Reserves.Count == 0)
                        paragraph.Add(new Chunk(string.Format("Tous "), _nameFont));
                    else
                    {
                        var i = 0;
                        foreach (var reserve in solo.Reserves)
                        {
                            paragraph.Add(GetUnitNameReferenceOrText(reserve.Key, reserve.Value, _nameFont));
                            ++i;
                            paragraph.Add(new Chunk(i < solo.Reserves.Count ? ", " : " ", _nameFont));
                        }
                    }
                    paragraph.Add(new Chunk(string.Format("{0}\n", coutSolo), _nameFont));

                    paragraph.Add(new Chunk(solo._Effet + Environment.NewLine, _textFont));

                    columnManager.AddUnbreakable(paragraph);
                }

                //columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);
                //columnManager.EndPage();

                paragraph = GetSubSectionTitleParagraph("Combattants affiliés\n");
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var u in GetUnitListForFaction(faction.Affilies))
                    columnManager.AddUnbreakable(u);
                columnManager.Render(2);

                //paragraph = GetSubSectionTitleParagraph("Combattants limités (+{0} PA pour la troupe, +{1} PA pour les Champions)\n", faction.CoutLimite, faction.CoutLimiteChampion);
                paragraph = GetSubSectionTitleParagraph("Combattants limités\n");
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var u in GetUnitListForFaction(faction.Limites, faction.CoutLimite, faction.CoutLimiteChampion))
                    columnManager.AddUnbreakable(u);
                columnManager.Render(2);

                paragraph = GetSubSectionTitleParagraph("Combattants interdits\n");
                columnManager.AddUnbreakable(paragraph);
                columnManager.Render(1);

                foreach (var u in GetUnitListForFaction(faction.Interdits))
                    columnManager.AddUnbreakable(u);
                columnManager.Render(2);
            }
            #endregion

        end:
            //CreateTableOfContents(doc, writer, 
            doc.Close();
        }

        private static Paragraph[] AddSeparatorParagraph(Paragraph[] paragraphs)
        {
            var lastParagraph = paragraphs.Last();
            var lineSeparator = new LineSeparator(1f, 100f, BaseColor.BLACK, Element.ALIGN_CENTER, -1);
            //if (lastParagraph.Alignment != Element.ALIGN_LEFT)
            //{
            //    var separatorParagraph = new Paragraph { lineSeparator };
            //    separatorParagraph.SpacingAfter = lastParagraph.SpacingAfter;
            //    lastParagraph.SpacingAfter = 0;
            //    var p = paragraphs.ToList();
            //    p.Add(separatorParagraph);
            //    return p.ToArray();
            //}
            lastParagraph.Add(lineSeparator);
            return paragraphs;
        }

        private static Paragraph[] GetUnitListForFaction(Dictionary<string, Unit> affilies, string paTroupe = null, string paChampion = null)
        {
            var result = new List<Paragraph>();
            //if (affilies.Count == 0)
            //{
            //    var paragraph = GetArmyPartParagraph();
            //    paragraph.Add(new Chunk("Aucun\n", _textFont));
            //    result.Add(paragraph);
            //}
            //else
            //{
            //    //var nonCharacterAffilies = affilies.Where(u => u.Value == null || !u.Value.IsPersonnage).OrderBy(u => u.Key).ToList();
            //    //var characterAffilies = affilies.Where(u => u.Value != null && u.Value.IsPersonnage).OrderBy(u => u.Key).ToList();

            //    var list = affilies.OrderBy(u => (u.Value == null || !u.Value.IsPersonnage ? "0" : "1") + u.Key).ToList();
            //    //var firstPersonnage = true;
            //    //var paragraph = new Paragraph { Leading = _textFont.Size };
            //    //result.Add(paragraph);

            //    var affiliesItems = new List<Paragraph>[2];
            //    affiliesItems[0] = new List<Paragraph>();
            //    affiliesItems[1] = new List<Paragraph>();
            //    //var maxWidths = new float[2];

            //    for (int i = 0; i < list.Count; i++)
            //    {
            //        var unit = list[i];

            //        //if (unit.Value.IsPersonnage && firstPersonnage)
            //        //{
            //        //    paragraph = new Paragraph { Leading = _textFont.Size };
            //        //    result.Add(paragraph);
            //        //    firstPersonnage = false;
            //        //}

            //        // try to regroup units with the same name and a different profile 
            //        // Eg. Fusilier (1), Fusilier (2) => Fusilier (1, 2)
            //        // or Esclave (Epée), Esclave (Lance) => Esclave (Epée, Lance)
            //        int j = 0, p;
            //        if ((p = unit.Key.LastIndexOf('(')) >= 0)
            //        {
            //            var prefix = unit.Key.Substring(0, p + 1);
            //            while (i + j + 1 < list.Count && list[i + j + 1].Key.LastIndexOf('(') >= 0 && list[i + j + 1].Key.StartsWith(prefix))
            //                ++j;
            //        }

            //        var index = unit.Value == null || !unit.Value.IsPersonnage ? 0 : 1;

            //        if (j == 0) // no regroupment
            //        {
            //            //paragraph = new Paragraph
            //            //{
            //            //    GetUnitNameReferenceOrText(unit.Key, unit.Value, _textFont)
            //            //};
            //            //paragraph.Add(GetUnitNameReferenceOrText(unit.Key, unit.Value, _textFont));
            //            //paragraph.Add(new Chunk(Environment.NewLine));
            //            affiliesItems[index].Add(new Paragraph(GetUnitNameReferenceOrText(unit.Key, unit.Value, _textFont)));
            //            //maxWidths[index] = Math.Max(maxWidths[index], _allTextFont.GetWidthPoint(t, _textFont.Size))
            //        }
            //        else
            //        {
            //            ++j;
            //            //paragraph = new Paragraph
            //            //{
            //            //    GetUnitNameReferenceOrText(unit.Key.Substring(0, unit.Key.LastIndexOf(')')), unit.Value, _textFont)
            //            //};
            //            //paragraph.Add(GetUnitNameReferenceOrText(unit.Key.Substring(0, unit.Key.LastIndexOf(')')), unit.Value, _textFont));
            //            var paragraph = new Paragraph(GetUnitNameReferenceOrText(unit.Key.Substring(0, unit.Key.LastIndexOf(')')), unit.Value, _textFont));
            //            for (var k = 1; k < j; ++k)
            //            {
            //                unit = list[i + k];
            //                p = unit.Key.LastIndexOf('(') + 1;
            //                var profileName = unit.Key.Substring(p, list[i + k].Key.Length - p - 1);
            //                var name = string.Format(", {0}{1}", profileName, k == j - 1 ? ")" : null);
            //                paragraph.Add(GetUnitNameReferenceOrText(name, unit.Value, _textFont));
            //            }
            //            i += j - 1;
            //            //paragraph.Add(new Chunk(Environment.NewLine));
            //            affiliesItems[unit.Value == null || !unit.Value.IsPersonnage ? 0 : 1].Add(paragraph);
            //        }
            //        //result.Add(paragraph);
            //    }

            //    if (paTroupe != null)
            //    {
            //        if (affiliesItems[0].Count == 0)
            //            affiliesItems[0].Add(new Paragraph("Aucune", _italicTextFont));
            //        if (affiliesItems[1].Count == 0)
            //            affiliesItems[1].Add(new Paragraph("Aucun", _italicTextFont));
            //    }

            //    //var column1WidthA = affiliesItems[0].Select(p => p.Aggregate("", (a, b) => a.ToString() + b.ToString()));
            //    //var dsdb = column1WidthA.Select(s => _allTextFont.GetWidthPoint(s, _textFont.Size)).ToList().Max();
            //    //_allTextFont.GetWidthPoint(t, _textFont.Size)

            //    var table = new PdfPTable(2) {WidthPercentage = 100};
            //    table.DefaultCell.BorderWidth = 0;
            //    //table.SetTotalWidth(new[] { 50f, 50f });

            //    if (paTroupe != null)
            //    {
            //        table.AddCell(new Phrase(string.Format("Troupe (+ {0} PA)", paTroupe), _nameFont));
            //        table.AddCell(new Phrase(string.Format("Champion (+ {0} PA)", paChampion), _nameFont));
            //    }

            //    var maxCount = Math.Max(affiliesItems[0].Count, affiliesItems[1].Count);
            //    for (int i = 0; i < maxCount; ++i)
            //    {
            //        if(i < affiliesItems[0].Count)
            //            table.AddCell(affiliesItems[0][i]);
            //        else
            //            table.AddCell(string.Empty);
            //        if (i < affiliesItems[1].Count)
            //            table.AddCell(affiliesItems[1][i]);
            //        else
            //            table.AddCell(string.Empty);
            //    }
            //    result.Add(new Paragraph {table});
            //}
            return result.ToArray();
        }

        public static Phrase GetUnitNameReferenceOrText(string nom, Unit unit, Font font)
        {
            return unit != null
                ? new Phrase { new Anchor(nom, font) { Reference = "#" + unit.FullName } }
                : new Phrase(nom, font);
        }

        public static Phrase GetUnitNameReference(Unit unit, Font font)
        {
            var unitName = new Anchor(unit.FullName, font) { Reference = "#" + unit.FullName };
            return new Phrase { unitName };
        }

        private static IEnumerable<string> Clean(this IEnumerable<string> items)
        {
            return items.Where(i => !string.IsNullOrWhiteSpace(i)).Select(i => i.Trim());
        }

        private static string JoinClean(string separator, IEnumerable<string> items)
        {
            return string.Join(separator, items.Clean());
        }

        private static string GetCoutText(string pa, string paChampion)
        {
            var pts = new[] { pa, paChampion }.Clean().ToList();
            return pts.Count == 0 ? string.Empty : string.Format("({0} PA)", string.Join("/", pts));
        }

        private static bool StartsWithVowel(this string s)
        {
            const string vowels = "aeiouyàèìòùéäëïöüÿ";
            if (s.Length <= 0)
                return false;
            var sl = s[0].ToString();
            return vowels.Any(v => string.Equals(v.ToString(), sl, StringComparison.InvariantCultureIgnoreCase));
        }

        private static Paragraph GetSectionTitleParagraph(string text, params object[] args)
        {
            Console.WriteLine(text);
            var phrase = new Phrase(string.Format(text, args), _sectionTitleFont);
            return new Paragraph(phrase) { FirstLineIndent = _sectionTitleIndent, SpacingAfter = 24, SpacingBefore = 36 };
        }

        private static Paragraph GetSectionTitleParagraphWithLink(int[] linkIndices, params string[] args)
        {
            var phrase = new Phrase();
            for (int i = 0; i < args.Length; ++i)
            {
                if (linkIndices.Contains(i))
                    phrase.Add(new Anchor(args[i], _sectionTitleFont) { Name = args[i] });
                else
                    phrase.Add(new Chunk(args[i], _sectionTitleFont));
            }
            return new Paragraph(phrase) { FirstLineIndent = _sectionTitleIndent, SpacingAfter = 24, SpacingBefore = 36 };
        }

        private static Paragraph GetSubSectionTitleParagraph(string text, params object[] args)
        {
            var phrase = new Phrase(string.Format(text, args), _subsectionTitleFont);
            return new Paragraph(phrase) { FirstLineIndent = _sectionTitleIndent, SpacingAfter = 12, SpacingBefore = 24 };
        }

        private static Paragraph AddParagraph(List<Paragraph> list, Paragraph newParagraph)
        {
            var lastParagraph = list[list.Count - 1];
            newParagraph.SpacingAfter = lastParagraph.SpacingAfter;
            lastParagraph.SpacingAfter = 0f;
            list.Add(newParagraph);
            return newParagraph;
        }

        //        private static Paragraph[] GetUnitParagraphs(Army army, Unit unit, float columnWidth)
        //        {
        //            var result = new List<Paragraph>();
        //            var unitParagraph = GetArmyPartParagraph();
        //            // create the bookmark for the unit
        //            var unitName = new Anchor(unit.FullName + Environment.NewLine, _nameFont) { Name = unit.FullName };
        //            unitParagraph.Add(new Phrase { unitName });

        //            result.Add(unitParagraph);

        //            AddParagraph(result, new Paragraph(new Chunk(string.Format("{0} PA\n", unit._points), _nameFont)) { });

        //            unitParagraph = AddParagraph(result, new Paragraph());

        //            var characValues = new List<string>
        //            {
        //                unit._M,
        //                unit._INI,
        //                string.Format("{0}/{1}", unit._ATT, unit._FOR),
        //                string.Format("{0}/{1}", unit._DEF, unit._RES),
        //                unit._TIR,
        //                unit._COU,
        //                unit._DIS
        //            };

        //            if (!string.IsNullOrWhiteSpace(unit._POU))
        //                characValues.Add(unit._POU);
        //            if (!string.IsNullOrWhiteSpace(unit._creation))
        //                characValues.Add(string.Format("{0}/{1}/{2}", unit._creation, unit._alteration, unit._destruction));

        //            var table = new PdfPTable(characValues.Count) { SpacingBefore = 8f, SpacingAfter = 8f };
        //            table.DefaultCell.BorderWidth = 1f;
        //            table.DefaultCell.BorderColor = BaseColor.WHITE;
        //            table.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
        //            table.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
        //            table.WidthPercentage = 100;
        //            table.DefaultCell.BackgroundColor = new BaseColor(240, 240, 240);

        //            int courage;

        //            var characNames = new List<string> { "M", "INI", "ATT/FOR", "DEF/RES", "TIR", Int32.TryParse(unit._COU, out courage) && courage < 0 ? "PEUR" : "COU", "DIS" };
        //            if (!string.IsNullOrWhiteSpace(unit._POU))
        //                characNames.Add("POU");
        //            if (!string.IsNullOrWhiteSpace(unit._creation))
        //                characNames.Add("C/A/D");

        //            var widths = characValues.Select((t, i) => Math.Max(_allTextFont.GetWidthPoint(t, _textFont.Size), _allTextFont.GetWidthPoint(characNames[i], _textFont.Size))).ToArray();
        //            var max = widths.Max();
        //            var spacingLeft = Math.Max(0, (columnWidth - widths.Sum()) / (characNames.Count - 2) - _cellSpace);
        //            widths = widths.Select(w => w == max ? w : w + spacingLeft).ToArray();

        //            table.SetTotalWidth(widths);

        //            foreach (var c in characNames)
        //                table.AddCell(new Phrase(c, _textFont));
        //            foreach (var c in characValues)
        //                table.AddCell(new Phrase(c, _textFont));
        //            unitParagraph.Add(table);

        //            if(!string.IsNullOrWhiteSpace(unit._equipment))
        //                AddParagraph(result, new Paragraph(new Chunk(string.Format("Équipement : {0}.\n", unit._equipment), _italicTextFont)));

        //            unitParagraph = AddParagraph(result, new Paragraph());
        //            if (!string.IsNullOrWhiteSpace(unit._competences))
        //            {
        //                var competences = unit._competences.Split('¤').ToList();
        //                if (competences.Count % 2 == 0)
        //                    throw new Exception("Erreur lors de l'analyse des compétences (le symbole '¤' doit être placé au début et à la fin du terme vers lequel le lien est créé : " + competences);

        //                for (int i = 0; i < competences.Count; i += 2)
        //                {
        //                    unitParagraph.Add(new Chunk(competences[i], _textFont));
        //                    if (i + 1 < competences.Count)
        //                    {
        //                        var anchor = new Anchor(competences[i + 1], _textFont) { Reference = "#" + competences[i + 1] };
        //                        unitParagraph.Add(anchor);
        //                    }
        //                }
        //                unitParagraph.Add(new Chunk(".\n", _textFont));
        //            }

        //            AddParagraph(result, new Paragraph(new Chunk(string.Format("{0} {1}. {2}.\n", unit.Rang, unit._Peuple, unit.Taille != Taille.Undefined ? unit.Taille.ToText() : unit._Taille), _italicTextFont)));
        //            if (unit._Special != null)
        //                unitParagraph.Add(new Chunk(string.Format("{0}\n", unit._Special), _specialFont));

        ////            result.Add(unitParagraph);


        //            if (unit.Special != null)
        //            {
        //                var specialParagraph = GetArmyPartParagraph(true);
        //                specialParagraph.Add(new Chunk(string.Format("{0}\n", unit.Special._nom), _specialFont));
        //                specialParagraph.Add(new Chunk(string.Format("{0}\n", unit.Special._Description), _textFont));
        //                result.Add(specialParagraph);
        //            }

        //            var artefacts = army.Artefacts.Where(a => a.Reserve == unit).ToList();
        //            result.AddRange(artefacts.Select(GetArtefactParagraph));
        //            return result.ToArray();
        //        }

        private static Paragraph[] GetUnitParagraphs(Army army, Unit unit, float columnWidth)
        {
            var result = new List<Paragraph>();
            var unitParagraph = GetArmyPartParagraph();
            // create the bookmark for the unit
            var unitName = new Anchor(unit.FullName + Environment.NewLine, _nameFont) { Name = unit.FullName };
            //unitParagraph.Add(new Phrase { unitName });

            var unitCost = string.Format("{0} PA\n", unit._points);

            //var table1 = new PdfPTable(2) { HorizontalAlignment = Element.ALIGN_LEFT };
            //table1.DefaultCell.BorderWidth = 0f;
            //table1.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            //table1.WidthPercentage = 98;

            //table1.AddCell(new PdfPCell(table1.DefaultCell) { Phrase = new Phrase { unitName } });
            //table1.AddCell(new PdfPCell(table1.DefaultCell) { HorizontalAlignment = Element.ALIGN_RIGHT, Phrase = new Phrase { new Chunk(unitCost, _nameFont) } });

            //var widths1 = (new[] { unit.FullName, unitCost }).Select(t => _allBoldTextFont.GetWidthPoint(t, _nameFont.Size)).ToArray();
            //table1.SetTotalWidth(widths1);

            //unitParagraph.Add(table1);
            unitParagraph.Add(new Phrase { unitName });

            int peur;
            Int32.TryParse(unit._PEU, out peur);

            var characValues = new List<TextPhrase>
            {
                new TextPhrase{ Text = unit._M, Phrase = new Phrase(unit._M, unit._concM == "0" ?  _textFont : _boldTextFont)},
                new TextPhrase{ Text = unit._INI, Phrase = new Phrase(unit._INI, unit._concINI == "0" ?  _textFont : _boldTextFont)},
                new TextPhrase{ Text = string.Format("{0}/{1}", unit._ATT, unit._FOR), Phrase = new Phrase { new Chunk(unit._ATT, unit._concATT == "0" ?  _textFont : _boldTextFont)} },
                new TextPhrase{ Text = string.Format("{0}/{1}", unit._DEF, unit._RES), Phrase = new Phrase { new Chunk(unit._DEF, unit._concDEF == "0" ?  _textFont : _boldTextFont)} },
                new TextPhrase{ Text = unit._TIR, Phrase = new Phrase(unit._TIR, unit._concTIR == "0" ?  _textFont : _boldTextFont)},
                new TextPhrase{ Text = peur > 0 ? unit._PEU : unit._COU, Phrase = new Phrase(peur > 0 ? unit._PEU : unit._COU, (peur > 0 ? unit._concPEU == "0" : unit._concCOU == "0") ? (peur > 0 ? _whiteTextFont : _textFont) : (peur > 0 ? _whiteBoldTextFont : _boldTextFont))},
                new TextPhrase{ Text = unit._DIS, Phrase = new Phrase(unit._DIS, unit._concDIS == "0" ?  _textFont : _boldTextFont)},
            };
            characValues[2].Phrase.Add(new Chunk("/", _textFont));
            characValues[2].Phrase.Add(new Chunk(unit._FOR, unit._concFOR == "0" ? _textFont : _boldTextFont));
            characValues[3].Phrase.Add(new Chunk("/", _textFont));
            characValues[3].Phrase.Add(new Chunk(unit._RES, unit._concRES == "0" ? _textFont : _boldTextFont));

            if (unit.IsMagical)
                characValues.Add(new TextPhrase { Text = unit._POU, Phrase = new Phrase(unit._POU, _textFont) });
            if (unit.IsDivine)
            {
                var text = string.Format("{0}/{1}/{2}", unit._creation, unit._alteration, unit._destruction);
                characValues.Add(new TextPhrase { Text = text, Phrase = new Phrase(text, _textFont) });
            }

            //foreach(var c in characteristics)
            //    unitParagraph.Add(new Chunk(c + Environment.NewLine, _textFont));

            var table = new PdfPTable(characValues.Count) { SpacingBefore = 8f, SpacingAfter = 8f };
            table.DefaultCell.BorderWidth = 1f;
            table.DefaultCell.BorderColor = BaseColor.WHITE;
            table.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            table.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            //table.DefaultCell.NoWrap = true;
            table.DefaultCell.PaddingBottom = 5f;
            table.WidthPercentage = 100;
            table.DefaultCell.BackgroundColor = new BaseColor(240, 240, 240);

            //var characNames = new List<string> { "M", "INI", "ATT/FOR", "DEF/RES", "TIR", peur > 0 ? "PEUR" : "COU", "DIS" };
            //if (unit.IsMagical)
            //    characNames.Add("POU");
            //if (unit.IsDivine)
            //    characNames.Add("C/A/D");

            var characNames = new List<TextPhrase>
            {
                new TextPhrase{ Text = "M", Phrase = new Phrase("M", unit._concM == "0" ?  _textFont : _boldTextFont)},
                new TextPhrase{ Text = "INI", Phrase = new Phrase("INI", unit._concINI == "0" ?  _textFont : _boldTextFont)},
                new TextPhrase{ Text = "ATT/FOR", Phrase = new Phrase { new Chunk("ATT", unit._concATT == "0" ?  _textFont : _boldTextFont)} },
                new TextPhrase{ Text = "DEF/RES", Phrase = new Phrase { new Chunk("DEF", unit._concDEF == "0" ?  _textFont : _boldTextFont)} },
                new TextPhrase{ Text = "TIR", Phrase = new Phrase("TIR", unit._concTIR == "0" ?  _textFont : _boldTextFont)},
                new TextPhrase{ Text = peur > 0 ? "PEU" : "COU", Phrase = new Phrase(peur > 0 ? "PEU" : "COU", (peur > 0 ? unit._concPEU == "0" : unit._concCOU == "0") ? (peur > 0 ? _whiteTextFont : _textFont) : (peur > 0 ? _whiteBoldTextFont : _boldTextFont))},
                new TextPhrase{ Text = "DIS", Phrase = new Phrase("DIS", unit._concDIS == "0" ?  _textFont : _boldTextFont)},
            };
            characNames[2].Phrase.Add(new Chunk("/", _textFont));
            characNames[2].Phrase.Add(new Chunk("FOR", unit._concFOR == "0" ? _textFont : _boldTextFont));
            characNames[3].Phrase.Add(new Chunk("/", _textFont));
            characNames[3].Phrase.Add(new Chunk("RES", unit._concRES == "0" ? _textFont : _boldTextFont));

            if (unit.IsMagical)
                characNames.Add(new TextPhrase { Text = "POU", Phrase = new Phrase("POU", _textFont) });
            if (unit.IsDivine)
            {
                characNames.Add(new TextPhrase { Text = "C/A/D", Phrase = new Phrase("C/A/D", _textFont) });
            }

            var widths = characValues.Select((t, i) => Math.Max(_allTextFont.GetWidthPoint(t.Text, _textFont.Size), _allTextFont.GetWidthPoint(characNames[i].Text, _textFont.Size))).ToArray();
            var max = widths.Max();
            var spacingLeft = Math.Max(0, (columnWidth - widths.Sum()) / (characNames.Count - 2) - _cellSpace);
            widths = widths.Select(w => w == max ? w : w + spacingLeft).ToArray();

            table.SetTotalWidth(widths);
            //table.LockedWidth = true;

            var odd = false;

            for (int i = 0; i < characNames.Count; i++)
            {
                var c = characNames[i];
                //var cell = new PdfPCell(new Phrase(c, _textFont))
                //var cell = new PdfPCell(table.DefaultCell)
                //{
                //    Phrase = new Phrase(c, _textFont),
                //    BackgroundColor = odd
                //            ? new BaseColor(240, 240, 240)
                //            : new BaseColor(232, 232, 232)
                //};
                //table.AddCell(cell);
                //odd = !odd;
                if (i == 5 && peur > 0)
                {
                    var cell = new PdfPCell(table.DefaultCell) { Phrase = c.Phrase };
                    cell.BackgroundColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }
                else
                    table.AddCell(c.Phrase);
            }
            odd = false;
            for (int i = 0; i < characValues.Count; i++)
            {
                var c = characValues[i];
                //var cell = new PdfPCell(new Phrase(c, _textFont))
                //var cell = new PdfPCell(table.DefaultCell)
                //{
                //    Phrase = new Phrase(c, _textFont),
                //    BackgroundColor = odd
                //            ? new BaseColor(248, 248, 248)
                //            : new BaseColor(240, 240, 240)
                //};
                //odd = !odd;
                //table.AddCell(cell);
                if (i == 5 && peur > 0)
                {
                    var cell = new PdfPCell(table.DefaultCell) { Phrase = c.Phrase };
                    cell.BackgroundColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }
                else
                    table.AddCell(c.Phrase);
            }
            unitParagraph.Add(table);

            if (unit.Equipment.Length > 0)
                unitParagraph.Add(new Chunk(string.Format("• {0}.\n", unit.Equipment), _italicTextFont));

            if (unit.Competences.Length > 0)
            {
                var competences = unit.Competences.Split('¤').ToList();
                if (competences.Count % 2 == 0)
                    throw new Exception("Erreur lors de l'analyse des compétences (le symbole '¤' doit être placé au début et à la fin du terme vers lequel le lien est créé : " + competences);

                for (int i = 0; i < competences.Count; i += 2)
                {
                    unitParagraph.Add(new Chunk(competences[i], _textFont));
                    if (i + 1 < competences.Count)
                    {
                        var anchor = new Anchor(competences[i + 1], _textFont) { Reference = "#" + competences[i + 1] };
                        unitParagraph.Add(anchor);
                    }
                }
                unitParagraph.Add(new Chunk(".\n", _textFont));
            }

            //if (!unit.IsMagical)
            //    unitParagraph.Add(new Chunk(string.Format("{0}.\n", unit._competences), _textFont));
            //else
            //{

            //var competences = unit._competences.Split(',').Select(c => c.Trim()).ToList();
            //for (int i = 0; i < competences.Count; ++i)
            //{
            //    var competence = competences[i];
            //    if ((competence.StartsWith("Initié") || competence.StartsWith("Adepte") || competence.StartsWith("Maître") || competence.StartsWith("Virtuose"))
            //        && competence.Contains('/'))
            //    {
            //        var compParts = competence.Split('/');
            //        unitParagraph.Add(new Chunk(compParts[0] + "/", _textFont));
            //        var paths = compParts[1].Split(',').Select(p => p.Trim()).ToList();
            //        for (int j = 0; j < paths.Count; j++)
            //        {
            //            var path = paths[j];
            //            var anchor = new Anchor(path, _textFont) { Reference = "#" + path };
            //            unitParagraph.Add(anchor);
            //            if (j < paths.Count - 1)
            //                unitParagraph.Add(new Chunk(", "));
            //        }
            //    }
            //    else
            //        unitParagraph.Add(new Chunk(string.Format("{0}{1}", competence, i < competences.Count - 1 ? ", " : null), _textFont));
            //}
            //}

            //unitParagraph.Add(new Chunk(string.Format("{0} {1}. {2}.\n", unit.Rank.Entity._name, unit.Peuple.Entity._name, unit.Size.Entity._name /*unit.Taille != Taille.Undefined ? unit.Taille.ToText() : unit._Taille*/), _italicTextFont));

            var table1 = new PdfPTable(2) { HorizontalAlignment = Element.ALIGN_LEFT };
            table1.DefaultCell.BorderWidth = 0f;
            table1.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            table1.WidthPercentage = 98;

            var peuple = unit.Peuple.Entity._qualifier;
            if (peuple.Contains("{Peuple}"))
            {
                string peupleRef = null;
                var suffixIndex = unit.FullName.IndexOf("(Dun-Scaîth)");
                if (suffixIndex >= 0)
                {
                    var refUnitName = unit.FullName.Substring(0, suffixIndex).Trim();
                    var refUnit = Army.AllArmies.Units.Values.FirstOrDefault(
                        u => (u.Peuple.Entity._name == "Keltois du clan des Drunes"
                            || u.Peuple.Entity._name == "Dévoreurs de Vile-Tys")
                            && u.FullName.StartsWith(refUnitName));
                    if (refUnit != null)
                        peupleRef = refUnit.Peuple.Entity._qualifier;
                }
                peuple = peuple.Replace("{Peuple}", peupleRef);
            }
            if (peuple.Contains("{Voie}"))
            {
                var voieIndex = unit.Competences.IndexOf("Être d");
                if (voieIndex >= 0)
                {
                    var voie = unit.Competences.Substring(voieIndex + 5);
                    if (voie.StartsWith("de Lumière", StringComparison.InvariantCultureIgnoreCase))
                        voie = "de la Lumière";
                    else if (voie.StartsWith("du Destin", StringComparison.InvariantCultureIgnoreCase))
                        voie = "du Destin";
                    else if (voie.StartsWith("des Ténèbres", StringComparison.InvariantCultureIgnoreCase))
                        voie = "des Ténèbres";
                    else
                        voie = null;
                    peuple = peuple.Replace("{Voie}", voie);
                }
            }

            var rankAndPeuple = string.Format("{0} {1}. {2}.\n", unit.Rank.Entity._name, peuple, unit.Size.Entity._name);
            if (unit.IsPersonnage)
                rankAndPeuple = "Champion " + rankAndPeuple;
            table1.AddCell(new PdfPCell(table1.DefaultCell) { Phrase = new Phrase(rankAndPeuple, _italicTextFont) });
            table1.AddCell(new PdfPCell(table1.DefaultCell) { HorizontalAlignment = Element.ALIGN_RIGHT, Phrase = new Phrase { new Chunk(unitCost, _nameFont) } });

            var widthRight = _allBoldTextFont.GetWidthPoint(unitCost, _nameFont.Size) + 10f;
            var widthLeft = columnWidth - widthRight;

            var widths1 = new[] { widthLeft, widthRight };
            table1.SetTotalWidth(widths1);

            unitParagraph.Add(table1);

            foreach (var capacity in unit.AssociatedCapacities)
                unitParagraph.Add(new Anchor(string.Format("{0}\n", capacity._name), _specialFont) { Reference = "#" + capacity._name });
            //            unitParagraph.Add(new Chunk(string.Format("{0} PA\n", unit._points), _nameFont));

            result.Add(unitParagraph);

            //var paragraph = new Paragraph(new Chunk(string.Format("{0} PA\n", unit._points), _nameFont)) { Alignment = Element.ALIGN_RIGHT };
            //result.Add(paragraph);
            //paragraph.SpacingAfter = unitParagraph.SpacingAfter;
            //unitParagraph.SpacingAfter = 0f;

            foreach (var capacity in unit.AssociatedCapacities)
            {
                var group = capacity.GroupedAssociatedUnits.First(g => g.Value.Contains(unit));
                if (group.Value.Count != 1 && group.Value.Count == unit.VariationCount && unit != group.Value.First())
                    continue;
                var specialParagraph = GetArmyPartParagraph(true);
                specialParagraph.Add(new Anchor(string.Format("{0}\n", capacity._name), _specialFont) { Name = capacity._name });
                //specialParagraph.Add(new Chunk(string.Format("{0}\n", capacity._description), _textFont));
                specialParagraph.AddRange(FormatDescription(capacity._description, _textFont));
                //specialParagraph.Add(new Chunk(Environment.NewLine, _textFont));
                result.Add(specialParagraph);
            }

            //var artefacts = army.Artifacts.Values.Where(a => a.Reserve == unit).ToList();
            result.AddRange(unit.AssociatedArtifacts
                .Where(artifact => artifact.GroupedAssociatedUnits.Count == 1)
                .Where(artifact =>
                    {
                        var group = artifact.GroupedAssociatedUnits.First();
                        return group.Value.Count == 1 || group.Value.Count != unit.VariationCount || unit == group.Value.First(); // describe the artifact only once for the first incarnation
                    })
                .Select(GetArtefactParagraph));
            return result.ToArray();
        }

        public static List<IElement> FormatDescription(string text, Font font)
        {
            var result = new List<IElement>();

            if (string.IsNullOrWhiteSpace(text))
                return result;
            text = text.Replace("•", "\n•");

            for (int i = 1; i <= 6; ++i)
                text = text.Replace(string.Format(".'{0}'", i), string.Format(".\n•'{0}'", i))
                    .Replace(string.Format(". '{0}'", i), string.Format(".\n•'{0}'", i));

            var lines = text.Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)).Select(l => l.Trim()).ToList();
            while (lines.Count > 0)
            {
                var builder = new StringBuilder();
                while (lines.Count > 0 && !lines[0].StartsWith("•"))
                {
                    builder.AppendLine(lines[0]);
                    lines.RemoveAt(0);
                }
                if (builder.Length > 0)
                    result.Add(new Chunk(builder.ToString(), font));

                if (lines.Count > 0 && lines[0].StartsWith("•"))
                {
                    var list = new List(List.UNORDERED, 10f) { IndentationLeft = 10f };
                    //list.SetListSymbol(lines[0].Length > 1 && lines[0][1] == '\'' ? string.Empty : "•");
                    list.SetListSymbol("•");
                    while (lines.Count > 0 && lines[0].StartsWith("•"))
                    {
                        list.Add(new ListItem(new Chunk(lines[0].Substring(1).Trim(), font)));
                        lines.RemoveAt(0);
                    }
                    result.Add(list);
                }
            }
            return result;
        }

        private static Paragraph GetArtefactParagraph(Artifact artifact)
        {
            var artifactParagraph = GetArmyPartParagraph(true);
            artifactParagraph.Add(new Chunk(string.Format("{0} ({1} PA)\n", artifact.FullName, artifact._points), _nameFont));

            artifactParagraph.Add(new Chunk(string.Format("{0}\n", artifact._effect), _textFont));
            if (artifact.GroupedAssociatedUnits.Count > 0)
            {
                artifactParagraph.Add(new Chunk("Réservé à ", _reservedToFont));
                bool isFirstGlobal = true;
                foreach (var unitGroup in artifact.GroupedAssociatedUnits)
                {
                    if (unitGroup.Value.Count == 1)
                    {
                        if (!isFirstGlobal)
                            artifactParagraph.Add(new Chunk(", ", _reservedToFont));
                        artifactParagraph.Add(GetUnitNameReferenceOrText(unitGroup.Value[0].FullName, unitGroup.Value[0], _reservedToFont));
                    }
                    else
                    {
                        if (!isFirstGlobal)
                            artifactParagraph.Add(new Chunk(", ", _reservedToFont));
                        var prefix = unitGroup.Key;
                        var prefixLength = prefix.Length;
                        if (unitGroup.Value.Count == unitGroup.Value[0].VariationCount && unitGroup.Value[0].IsPersonnage) // all variations of the character are there
                        {
                            if (!isFirstGlobal)
                                artifactParagraph.Add(new Chunk(", ", _reservedToFont));
                            artifactParagraph.Add(GetUnitNameReferenceOrText(prefix.Substring(0, prefixLength - 2), unitGroup.Value[0], _reservedToFont));
                        }
                        else
                        {
                            var isFirst = true;
                            foreach (var artifactUnit in unitGroup.Value)
                            {
                                if (!isFirst)
                                    artifactParagraph.Add(new Chunk(", ", _reservedToFont));
                                else
                                    isFirst = false;
                                var linkText = artifactUnit.FullName.Substring(prefixLength).Replace(")", null).Replace("  ", " ");
                                artifactParagraph.Add(GetUnitNameReferenceOrText(prefix + linkText, artifactUnit, _reservedToFont));
                                prefix = null;
                            }
                            artifactParagraph.Add(new Chunk(")", _reservedToFont));
                        }
                    }
                    isFirstGlobal = false;
                }
                artifactParagraph.Add(new Chunk(".\n", _reservedToFont));
            }

            //if (!string.IsNullOrWhiteSpace(artifact._Reserve))
            //{
            //    artifactParagraph.Add(new Chunk(artifact.Reserve != null ? "Réservé à " : "Réservé aux ", _reservedToFont));
            //    artifactParagraph.Add(GetUnitNameReferenceOrText(artifact._Reserve, artifact.Reserve, _reservedToFont));
            //    artifactParagraph.Add(new Chunk(Environment.NewLine, _reservedToFont));
            //}
            return artifactParagraph;
        }

        private static string NonZeroOrX(string value)
        {
            return value == "0" ? null : (value == "99" ? "X" : value);
        }

        private static Paragraph GetSpellParagraph(Spell spell, float columnWidth)
        {
            var spellParagraph = GetArmyPartParagraph(true);
            spellParagraph.Add(new Chunk(string.Format("{0}\n", spell._name), _nameFont));
            var composants = new Dictionary<string, string> 
            { 
                { "Air", NonZeroOrX(spell._air) },
                { "Eau", NonZeroOrX(spell._water) },
                { "Feu", NonZeroOrX(spell._fire) },
                { "Terre", NonZeroOrX(spell._earth) },
                { "Lumière", NonZeroOrX(spell._light) },
                { "Ténèbres", NonZeroOrX(spell._darkness) },
            };

            var composantsText = string.Join(", ", composants.Where(kvp => !string.IsNullOrWhiteSpace(kvp.Value)).Select(kvp => string.Format("{0} ({1})", kvp.Key, kvp.Value)).ToArray());
            var voieText = string.Format("Voie : {0}", string.Join(", ", spell.Voies));
            var difficulteText = string.Format("Difficulté : {0}", ValueOrX(spell._difficulty));
            var porteeText = string.Format("Portée : {0}", spell._range);
            var frequenceText = string.Format("Fréquence : {0}", ValueOrX(spell._frequency));
            var puissanceText = string.Format("Puissance : {0}", ValueOrX(spell._power));
            var aireEffetText = string.Format("Aire d'effet : {0}", spell._aoe);
            var dureeText = string.Format("Durée : {0}", spell._duration);

            var table = new PdfPTable(2);
            table.DefaultCell.BorderWidth = 0f;
            table.WidthPercentage = 100;
            //var cell = new PdfPCell(new Phrase("Header spanning 3 columns"));
            //cell.Colspan = 3;
            //cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
            //table.AddCell(cell);
            //table.AddCell(new PdfPCell(new Phrase(voieText, _textFont)) { Colspan = 2, BorderWidth = 0f});
            //table.AddCell(new PdfPCell(new Phrase(composantsText, _textFont)) { Colspan = 2, BorderWidth = 0f });

            var leftColumn = new[] { voieText, difficulteText, aireEffetText, dureeText };
            var rightColumn = new[] { composantsText, porteeText, frequenceText, puissanceText };

            var leftColumnWidth = leftColumn.Select(t => _allTextFont.GetWidthPoint(t, _textFont.Size)).Max();
            var rightColumnWidth = rightColumn.Select(t => _allTextFont.GetWidthPoint(t, _textFont.Size)).Max();

            leftColumnWidth = Math.Min(leftColumnWidth, columnWidth - rightColumnWidth - _cellSpace);

            table.SetWidths(new[] { leftColumnWidth, rightColumnWidth });

            for (int i = 0; i < leftColumn.Length; ++i)
            {
                table.AddCell(new Phrase(leftColumn[i], _italicTextFont));
                table.AddCell(new Phrase(rightColumn[i], _italicTextFont));
            }

            //result.Add(table);
            //sortilegeParagraph.Add(new Chunk(string.Join(" ; ", new[] { composantsText, voieText, puissanceText, frequenceText }) + Environment.NewLine, _textFont));


            //sortilegeParagraph.Add(new Chunk(string.Join(" ; ", new[] { difficulteText, porteeText, aireEffetText, dureeText }) + Environment.NewLine, _textFont));

            //table = new PdfPTable(2);
            //table.DefaultCell.BorderWidth = 0f;
            //var cell = new PdfPCell(new Phrase("Header spanning 3 columns"));
            //cell.Colspan = 3;
            //cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
            //table.AddCell(cell);
            //table.AddCell(new PdfPCell(new Phrase(aireEffetText, _textFont)) { Colspan = 2, BorderWidth = 0f });
            //table.AddCell(new PdfPCell(new Phrase(dureeText, _textFont)) { Colspan = 2, BorderWidth = 0f });  
            //result.Add(table);
            spellParagraph.Add(table);

            //sortilegeParagraph = GetArmyPartParagraph(true);
            //spellParagraph.Add(new Chunk(string.Format("\n{0}\n", spell._effect), _textFont));
            spellParagraph.Add(new Chunk(Environment.NewLine, _textFont));
            spellParagraph.AddRange(FormatDescription(spell._effect, _textFont));

            return spellParagraph;
            //result.Add(sortilegeParagraph);
            //return result.ToArray();
        }

        private static string ValueOrX(string value)
        {
            return value == "99" ? "X" : value;
        }

        private static Paragraph GetMiracleParagraph(Miracle miracle, float columnWidth)
        {
            var miracleParagraph = GetArmyPartParagraph(true);
            miracleParagraph.Add(new Chunk(string.Format("{0}\n", miracle._name), _nameFont));

            var requirementText = string.Format("C/A/D: {0}/{1}/{2}", miracle._creation, miracle._alteration, miracle._destruction);
            var culteText = string.Format("Culte : {0}", string.Join(", ", miracle.Cults));
            var difficulteText = string.Format("Difficulté : {0}", ValueOrX(miracle._difficulty));
            var porteeText = string.Format("Portée : {0}", miracle._range);
            var aireEffetText = string.Format("Aire d'effet : {0}", miracle._aoe);
            var dureeText = string.Format("Durée : {0}", miracle._duration);
            var fervorText = string.Format("Ferveur : {0}", ValueOrX(miracle._fervor));

            var table = new PdfPTable(2);
            table.DefaultCell.BorderWidth = 0f;
            table.WidthPercentage = 100;

            var leftColumn = new[] { culteText, difficulteText, aireEffetText, dureeText };
            var rightColumn = new[] { requirementText, fervorText, porteeText, string.Empty };

            var leftColumnWidth = leftColumn.Select(t => _allTextFont.GetWidthPoint(t, _textFont.Size)).Max();
            var rightColumnWidth = rightColumn.Select(t => _allTextFont.GetWidthPoint(t, _textFont.Size)).Max();

            leftColumnWidth = Math.Min(leftColumnWidth, columnWidth - rightColumnWidth - _cellSpace);

            table.SetWidths(new[] { leftColumnWidth, rightColumnWidth });

            for (int i = 0; i < leftColumn.Length; ++i)
            {
                table.AddCell(new Phrase(leftColumn[i], _italicTextFont));
                table.AddCell(new Phrase(rightColumn[i], _italicTextFont));
            }

            miracleParagraph.Add(table);
            //miracleParagraph.Add(new Chunk(string.Format("\n{0}\n", miracle._effect), _textFont));
            miracleParagraph.Add(new Chunk(Environment.NewLine, _textFont));
            miracleParagraph.AddRange(FormatDescription(miracle._effect, _textFont));

            return miracleParagraph;
        }

        private static Paragraph GetArmyPartParagraph(bool justified = false)
        {
            var result = new Paragraph { SpacingAfter = 16, Leading = _textFont.Size * 1.2f };
            if (justified)
                result.Alignment = Element.ALIGN_JUSTIFIED;
            return result;
        }

        //private static void EndPage()
        //{
        //    _newPageRequired = true;
        //}

        //private static float _columnTop;
        //private static bool _newPageRequired;

        //private static void RenderColumns(Document doc, ColumnText ct, int columnCount)
        //{
        //    if (_newPageRequired)
        //    {
        //        _columnTop = doc.Top;
        //        doc.NewPage();
        //        _newPageRequired = false;
        //    }
        //    else if (_columnTop <= doc.Bottom)
        //        _columnTop = doc.Top;

        //    //float[] left = { doc.Left + 90f , doc.Top - 80f,
        //    //      doc.Left + 90f, doc.Top - 170f,
        //    //      doc.Left, doc.Top - 170f,
        //    //      doc.Left , doc.Bottom };

        //    //float[] right = { doc.Left + colwidth, doc.Top - 80f,
        //    //        doc.Left + colwidth, doc.Bottom };

        //    //float[] left =
        //    //    {
        //    //        doc.Left, doc.Top - 80f,
        //    //        doc.Left, doc.Bottom
        //    //    };

        //    //float[] right =
        //    //    {
        //    //        doc.Left + colwidth, doc.Top - 80f,
        //    //        doc.Left + colwidth, doc.Bottom
        //    //    };

        //    //float[] left2 =
        //    //    {
        //    //        doc.Right - colwidth, doc.Top - 80f,
        //    //        doc.Right - colwidth, doc.Bottom
        //    //    };

        //    //float[] right2 =
        //    //    {
        //    //        doc.Right, doc.Top - 80f,
        //    //        doc.Right, doc.Bottom
        //    //    };

        //    int status = 0;
        //    float currentBottom = _columnTop - 10f;
        //    //Checks the value of status to determine if there is more text
        //    //If there is, status is 2, which is the value of NO_MORE_COLUMN

        //    bool requiresMoreSpace = true;
        //    while (currentBottom > doc.Bottom && requiresMoreSpace)
        //    {
        //        var ctCopy = ColumnText.Duplicate(ct);
        //        //left[3] = right[3] = left2[3] = right2[3] = currentBottom;
        //        requiresMoreSpace = TryRenderColumns(doc, true, ctCopy, status, currentBottom, columnCount);
        //        if (requiresMoreSpace)
        //            currentBottom = Math.Max(doc.Bottom, currentBottom - 10f);
        //    }
        //    TryRenderColumns(doc, false, ct, status, currentBottom, columnCount);
        //    _columnTop = currentBottom;
        //}

        //private static bool TryRenderColumns(Document doc, bool simulate, ColumnText ct, int status, float currentBotton, int columnCount)
        //{
        //    float gutter = 15f;
        //    float colwidth = (doc.Right - doc.Left - gutter * (columnCount - 1)) / columnCount;

        //    var i = 0;
        //    while (ColumnText.HasMoreText(status))
        //    {
        //        var left = doc.Left + i * (gutter + colwidth);
        //        var right = left + colwidth;
        //        float[] leftBorder = { left, _columnTop, left, currentBotton };
        //        float[] rightBorder = { right, _columnTop, right, currentBotton };

        //        //Writing the column
        //        ct.SetColumns(leftBorder, rightBorder);

        //        //Needs to be here to prevent app from hanging
        //        ct.YLine = _columnTop;
        //        //Commit the content of the ColumnText to the document
        //        //ColumnText.Go() returns NO_MORE_TEXT (1) and/or NO_MORE_COLUMN (2)
        //        //In other words, it fills the column until it has either run out of column, or text, or both
        //        status = ct.Go(simulate);

        //        if (status == ColumnText.NO_MORE_COLUMN && i == columnCount - 1)
        //        {
        //            if (simulate)
        //                return true;
        //            i = 0;
        //            doc.NewPage();
        //        }
        //        else
        //            ++i;
        //    }
        //    return false;
        //}

        public static void CreateTableOfContents(Document doc, PdfWriter writer, int pageCount, PdfPTable tocTable, List<TableOfContentsEntry> toc)
        {
            doc.NewPage();
            pageCount++;

            doc.Add(new Paragraph("Table of Contents", FontFactory.GetFont("Arial", 18, Font.BOLD)));
            doc.Add(new Chunk(Environment.NewLine));

            tocTable = new PdfPTable(2);

            foreach (TableOfContentsEntry content in toc)
            {
                var nameCell = new PdfPCell(tocTable)
                {
                    Border = Rectangle.NO_BORDER,
                    Padding = 6f,
                    Phrase = new Phrase(content.Title)
                };
                tocTable.AddCell(nameCell);

                var pageCell = new PdfPCell(tocTable)
                {
                    Border = Rectangle.NO_BORDER,
                    Padding = 6f,
                    Phrase = new Phrase(content.Page)
                };
                tocTable.AddCell(pageCell);
            }

            doc.Add(tocTable);
            doc.Add(new Chunk(Environment.NewLine));

            /** Reorder pages so that TOC will will be the second page in the doc
            * right after the title page**/
            int tocPage = writer.PageNumber - 1;
            int total = writer.ReorderPages(null);
            var order = new int[total];

            for (int i = 0; i < total; ++i)
            {
                if (i == 0)
                {
                    order[i] = 1;
                }
                else if (i == 1)
                {
                    order[i] = tocPage;
                }
                else
                {
                    order[i] = i;
                }
            }

            writer.ReorderPages(order);
        }
    }

    public class TextPhrase
    {
        public string Text;
        public Phrase Phrase;
    }

    public struct TableOfContentsEntry
    {
        private string _title;
        private string _page;

        public TableOfContentsEntry(string title, string page)
        {
            _title = title;
            _page = page;
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; }
        }

        public string Page
        {
            get { return _page; }
            set { _page = value; }
        }
    }
}