﻿using System;
using System.Security.AccessControl;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Security.Cryptography.X509Certificates;
using DocumentFormat.OpenXml.Vml;
using w14 = DocumentFormat.OpenXml.Office2010.Word;

namespace StudyOpenXml
{
    public class DocManipulation
    {
        public DocManipulation()
        {
        }


        public static void createDocument(string filepath, string styleid, string stylename)
        {
            //Path:  /Users/fabianvalverde/Documents/GitHub/StudyOpenXml

            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Insert other code here

                //This piece is adding the main part of the document
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                //Then, this is creating the structure to manipulate every part of the document
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                ParagraphProperties paraProperties = new ParagraphProperties();
                StyleDefinitionsPart styleDefinition = wordDocument.MainDocumentPart.StyleDefinitionsPart;
                Run run = para.AppendChild(new Run());

                /*Prefer this way:

                 Document document = new Document();
                 Body body = new Body();
                 Paragraph para = new Paragraph();
                
                 Then:
                 document.append(body);
                 body.append(para);
                 para.append(run);
                
                 Really easier to understand this way, much less complex to read*/

                run.AppendChild(new Text("My first text"));

                //Just an experiment to create an style doc
                //Have to study this code and below methods better
                para.PrependChild(paraProperties);
                styleDefinition = AddStylesPartToPackage(wordDocument);
                AddNewStyle(styleDefinition, styleid, stylename);
                paraProperties.ParagraphStyleId = new ParagraphStyleId { Val = styleid };
            }
        }

        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            //Adding a part of type Style as a child of the main document
            StyleDefinitionsPart part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            //Styles root contains all the styles and saves the part
            Styles root = new Styles();
            //I think this is to saves the Model with the new style definition, not sure.
            root.Save(part);
            return part;
        }

        // Return true if the style id is in the document
        public static bool IsStyleIdInDocument(StyleDefinitionsPart styleDefinitionsPart, string styleid)
        {
            // Get access to the Styles element for this document (directly the styles).
            Styles s = styleDefinitionsPart.Styles;

            // Check that there are styles and how many (in the main document)
            int n = s.Elements<Style>().Count();
            if (n == 0)
            {
                return false;
            }


            // Look for a match on styleid.
            //Where the style element in our main document matchs the id and the type Paragraph
            Style style = s.Elements<Style>()
                .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                .FirstOrDefault();
            if (style == null)
                return false;

            return true;
        }

        private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart,
        string styleid, string stylename)
        {
            // Get access to the root element of the styles part in the main document.
            Styles styles = styleDefinitionsPart.Styles;

            // Create a new paragraph style and specify some of the properties.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true
            };


            StyleName styleName1 = new StyleName() { Val = stylename };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            style.Append(styleName1);
            style.Append(basedOn1);
            style.Append(nextParagraphStyle1);

            // Create the StyleRunProperties object and specify some of the run properties.
            //This is for the text
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
            Italic italic1 = new Italic();
            // Specify a 12 point size.
            FontSize fontSize1 = new FontSize() { Val = "24" };
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        }

        // Return styleid that matches the styleName, or null when there's no match.
        public static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
        {
            //Remember the style definitions part contains every style.
            StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                .Where(s => s.Val.Value.Equals(styleName) &&
                    (((Style)s.Parent).Type == StyleValues.Paragraph))
                .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
        }

        //THIS IS FOR TABLE!!!--------------------------------------------------------------------------------------------


        // Insert a table into a word processing document.
        public static void createTable(string filepath)
        {

            /*This is the structure we want to recreate
 *    < w:document xmlns:w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
        < w:body >
            < w:tbl >
                < w:tblGrid >
                    < w:gridCol />
                    < w:gridCol />
                </ w:tblGrid >
                < w:tblPr >
                    < w:tblBorders >
                        <w:top/>
                        <w:bottom/>
                        <w:right/>
                        <w:left/>
                        <w:insideH/>
                        <w:insideV/>
                    </ w:tblBorders >
                </ w:tblPr >
                < w:tr >
                    < w:tc >
                        < w:p >
                            < w:r >
                                < w:t > Working </ w:t >
                            </ w:r >
                        </ w:p >
                        < w:tcPr >
                        </ w:tcPr >
                    </ w:tc >
                </ w:tr >
            </ w:tbl >
        </ w:body >*/

            using (WordprocessingDocument wordDocument =
                        WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                //creating the doc
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();//<w:document>
                var body = mainPart.Document.AppendChild(new Body());//<w:body>
                var table = new Table();//<w:tbl>
                var tableGrid = new TableGrid();//<w:tblGrid>
                var gridCol1 = new GridColumn() { };//<w:gridCol/>
                var gridCol2 = new GridColumn();//<w:gridCol/>
                var tblPr = new TableProperties();//<w:tblPr>
                var width = new TableWidth() { Width = "100", Type = TableWidthUnitValues.Pct};//<w:tblW/>
                var tblBorder = new TableBorders();//<w:tblBorders>


                //Table borders
                var borderColor = "FF8000";

                var topBorder = new TopBorder();//<w:top/>
                topBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                topBorder.Size = 8;
                topBorder.Color = borderColor;

                var bottomBorder = new BottomBorder();//<w:bottom/>
                bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                bottomBorder.Size = 8;
                bottomBorder.Color = borderColor;

                var rightBorder = new RightBorder();//< w:right />
                rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                rightBorder.Size = 8;
                rightBorder.Color = borderColor;

                var leftBorder = new LeftBorder();//<w:left/>
                leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                leftBorder.Size = 8;
                leftBorder.Color = borderColor;

                var insideHorizontalBorder = new InsideHorizontalBorder();//<w:insideH/>
                insideHorizontalBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                insideHorizontalBorder.Size = 8;
                insideHorizontalBorder.Color = borderColor;

                var insideVerticalBorder = new InsideVerticalBorder();//<w:insideV/>
                insideVerticalBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                insideVerticalBorder.Size = 8;
                insideVerticalBorder.Color = borderColor;

               //tblBorder's child

/*             < w:tblPr >
                   < w:tblBorders >
                        < w:top />
                        < w:bottom />
                        < w:right />
                        < w:left />
                        < w:insideH />
                        < w:insideV />
                   </ w:tblBorders >
                </ w:tblPr >               */

                tblPr.AppendChild(tblBorder);

                tblBorder.AppendChild(topBorder);
                tblBorder.AppendChild(bottomBorder);
                tblBorder.AppendChild(rightBorder);
                tblBorder.AppendChild(leftBorder);
                tblBorder.AppendChild(insideHorizontalBorder);
                tblBorder.AppendChild(insideVerticalBorder);



                //-----------------------------------------------------------------------------------------




                //tableGrid's child

                /*                < w:tblGrid >
                                     < w:gridCol />
                                     < w:gridCol />
                                  </ w:tblGrid >     */
                table.AppendChild(tblPr);//You'll see tblPr before tblGrid

                tableGrid.AppendChild(gridCol1);
                tableGrid.AppendChild(gridCol2);

                //If we check the structure we're following we can see table is parent of tblgrid, tblPr and tblRow
                table.AppendChild(tableGrid);

                

                var para = new Paragraph(new Run(new Text("Cell1")));
                var para2 = new Paragraph(new Run(new Text("Cell2")));

                //Creating content
                var row1 = new TableRow();//<w:tr>

                var cell1 = new TableCell();//<w:tc>
                var cellProp = new TableCellProperties();//< w:tcPr >
                var cellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4000" };
                cell1.AppendChild(para);
                cellProp.AppendChild(cellWidth);
                cell1.AppendChild(cellProp);
                
                

                var cell2 = new TableCell();//<w:tc>
                var cellProp2 = new TableCellProperties();//<w:tcPr>
                var cellWidth2 = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1000" };
                cell2.AppendChild(para2);
                cell2.AppendChild(cellProp2);
                cellProp2.AppendChild(cellWidth2);


                //Here's where the order matters
                //Where you're appending childs to the same parent
                row1.Append(cell2);
                row1.Append(cell1);
                

                table.AppendChild(row1);
                table.AppendChild(width);//We'll see <w:tblW> child as the last one

                body.Append(table);
            }
        }

        public static void createCheckBox(string filepath, string internalName, int internalId, string textAfterTextbox)
        {
            using (WordprocessingDocument wordDocument =
             WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                //creating the doc
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                var run1 = new Run(
                                new FieldChar(
                                    new FormFieldData(
                                        new FormFieldName() { Val = internalName },
                                        new Enabled(),
                                        new CalculateOnExit() { Val = OnOffValue.FromBoolean(false) },
                                        new CheckBox(
                                            new AutomaticallySizeFormField(),
                                            new DefaultCheckBoxFormFieldState() { Val = OnOffValue.FromBoolean(false) }))
                                )
            {
                FieldCharType = FieldCharValues.Begin
            }
        );
                var run2 = new Run(new FieldCode(" FORMCHECKBOX ") { Space = SpaceProcessingModeValues.Preserve });
                var run3 = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
                var run4 = new Run(new Text(textAfterTextbox));

                Paragraph para = new Paragraph(
                        run1,
                        new BookmarkStart() { Name = internalName, Id = new StringValue(internalId.ToString()) },
                        run2,
                        run3,
                        new BookmarkEnd() { Id = new StringValue(internalId.ToString()) },
                        run4
                    );
                body.AppendChild(para);

            }
        }

        public static void createCheckBox2(string filepath)
        {

            using (WordprocessingDocument wordDocument =
 WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                //creating the doc
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());
                Paragraph paragraph = new Paragraph();

                //SdtBlock block = new SdtBlock(); check google
                SdtRun sdt = new SdtRun();//<w:sdt>
                SdtProperties sdtPr = new SdtProperties();//<w:sdtPr>
                SdtId id = new SdtId() { Val = -15934659 };//<w:id>
                w14.SdtContentCheckBox checkbox = new w14.SdtContentCheckBox();//<w14:checkbox>
                w14.Checked checkedd = new w14.Checked() { Val = w14.OnOffValues.Zero };//<w14:checked>
                w14.CheckedState checkedState = new w14.CheckedState() { Font = "MS Gothic", Val = "2612" };//<w14:checkedState>
                w14.UncheckedState uncheckedState = new w14.UncheckedState() { Font = "MS Gothic", Val = "2610" };//<w14:uncheckedState>

                sdt.Append(sdtPr);

                /*<w:sdtPr>
                    <w:id w:val="-15934659"/>
                    <w14:checkbox>
                        <w14:checked w14:val="1"/>
                        <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
                        <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
                    </w14:checkbox>
                </w:sdtPr>*/

                //id and checkbox are sdtPr's childs
                sdtPr.Append(id);
                sdtPr.Append(checkbox);
                //and below we can see checkbox's childs
                checkbox.Append(checkedd);
                checkbox.Append(checkedState);
                checkbox.Append(uncheckedState);

                //----------------------------------------------------------------------

                SdtContentRun sdtContentRun = new SdtContentRun();//<w:sdtContent>

                Run run = new Run();//<w:r>
                RunProperties runProperties = new RunProperties();//<w:rPr>
                RunFonts runFonts = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "MS Gothic", HighAnsi = "MS Gothic", EastAsia = "MS Gothic" };//<w:rFonts>

                runProperties.Append(runFonts);
                Text text1 = new Text();//<w:t>
                text1.Text = "☐";

                run.Append(runProperties);
                run.Append(text1);

                sdtContentRun.Append(run);

                //-----------------------------------------------------------------------

                ProofError spellStart = new ProofError() { Type = ProofingErrorValues.SpellStart };//<w:proofErr>

                ProofError spellEnd = new ProofError() { Type = ProofingErrorValues.SpellEnd };//<w:proofErr>

                sdt.Append(sdtContentRun);

                paragraph.AppendChild(sdt);

                //Remember the order we append childs of the same element is important (Paragraph appends)
                /*<w:proofErr w:type="spellStart"/>
                    <w:r>
                        <w:t>Dgdsg</w:t>
                    </w:r>
                <w:proofErr w:type="spellEnd"/>*/

                paragraph.AppendChild(spellStart);

                Run run2 = new Run();
                Text text2 = new Text();
                text2.Text = "This is the text";
                run2.AppendChild(text2);
                paragraph.AppendChild(run2);

                paragraph.AppendChild(spellEnd);

                body.AppendChild(paragraph);

            }
        }

        public static void changeBackgroundTable(string filepath)
        {
            using (WordprocessingDocument wordDocument =
WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {

                //creating the doc
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());
                var table = new Table();//<w:tbl>
                var tblPr = new TableProperties();//<w:tblPr>
                var tblStyle = new TableStyle() { Val = "TableGrid"};
                var width = new TableWidth() { Width = "100", Type = TableWidthUnitValues.Pct };//<w:tblW/>
                var tblBorder = new TableBorders();//<w:tblBorders>
                var tblLook = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, 
                    LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

                var tableGrid = new TableGrid();//<w:tblGrid>
                var gridCol1 = new GridColumn();//<w:gridCol/>

                var borderColor = "A5A5A5";

                var topBorder = new TopBorder();//<w:top/>
                topBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
                topBorder.Size = 3;
                topBorder.Color = borderColor;

                var bottomBorder = new BottomBorder();//<w:bottom/>
                bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
                bottomBorder.Size = 3;
                bottomBorder.Color = borderColor;

                var rightBorder = new RightBorder();//< w:right />
                rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
                rightBorder.Size = 3;
                rightBorder.Color = borderColor;

                var leftBorder = new LeftBorder();//<w:left/>
                leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
                leftBorder.Size = 3;
                leftBorder.Color = borderColor;

                //-----------------------------------------------------------------------------

                var tr = new TableRow();
                var tc = new TableCell();
                var tcPr = new TableCellProperties();
                var tcW = new TableCellWidth();
                var shd = new Shading();

                var p = new Paragraph();

                var r = new Run();

                ProofError spellStart = new ProofError() { Type = ProofingErrorValues.SpellStart };//<w:proofErr>

                ProofError spellEnd = new ProofError() { Type = ProofingErrorValues.SpellEnd };//<w:proofErr>
            }
        }

        public static void highlightText(string filepath)
        {
            using (WordprocessingDocument wordDocument =
WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());
                var p = new Paragraph();

                var pPr = new ParagraphProperties();
                var rPr = new RunProperties();
                var color = new Color();

                var run = new Run();

                var highLight = new Highlight();

                var sectPr = new SectionProperties();

                var pgSz = new PageSize();
                var pgMar = new PageMargin();
                var cols = new Columns();
                var docGrid = new DocGrid();


            }
        }
    }
}
