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
using static Stylexml.StyleManipulation;
using w14 = DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Net.Http;

namespace Documentxml
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
                AddNewTableStyle(styleDefinition, styleid, stylename);
                paraProperties.ParagraphStyleId = new ParagraphStyleId { Val = styleid };
            }
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
                var width = new TableWidth() { Width = "100", Type = TableWidthUnitValues.Pct };//<w:tblW/>
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

                //SdtBlock block = new SdtBlock();
                SdtRun sdt = new SdtRun();//<w:sdt>
                SdtProperties sdtPr = new SdtProperties();//<w:sdtPr>
                //SdtId id = new SdtId() { Val = -15934659 };//<w:id>
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
                //sdtPr.Append(id);
                sdtPr.Append(checkbox);
                //and below we can see checkbox's childs
                checkbox.Append(checkedd);
                checkbox.Append(checkedState);
                checkbox.Append(uncheckedState);

                //----------------------------------------------------------------------

                SdtContentRun sdtContentRun = new SdtContentRun();//<w:sdtContent>

                Run run = new Run();//<w:r>
                /*                RunProperties runProperties = new RunProperties();//<w:rPr>
                                RunFonts runFonts = new RunFonts()
                                {
                                    Hint = FontTypeHintValues.EastAsia,
                                    Ascii = "MS Gothic",
                                    HighAnsi = "MS Gothic",
                                    EastAsia = "MS Gothic"
                                };*/

                //runProperties.Append(runFonts);
                Text text1 = new Text();//<w:t>
                text1.Text = "☐";
                //☒☐

                //run.Append(runProperties);
                run.Append(text1);

                sdtContentRun.Append(run);

                //-----------------------------------------------------------------------

                //ProofError spellStart = new ProofError() { Type = ProofingErrorValues.SpellStart };//<w:proofErr>

                //ProofError spellEnd = new ProofError() { Type = ProofingErrorValues.SpellEnd };//<w:proofErr>

                sdt.Append(sdtContentRun);

                paragraph.AppendChild(sdt);

                //Remember the order we append childs of the same element is important (Paragraph appends)
                /*<w:proofErr w:type="spellStart"/>
                    <w:r>
                        <w:t>Dgdsg</w:t>
                    </w:r>
                <w:proofErr w:type="spellEnd"/>*/

                //paragraph.AppendChild(spellStart);

                Run run2 = new Run();
                Text text2 = new Text();
                text2.Text = "This is the text";
                run2.AppendChild(text2);
                paragraph.AppendChild(run2);

                //paragraph.AppendChild(spellEnd);

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
                var table = body.AppendChild(new Table());
                var tblPr = table.AppendChild(new TableProperties(new TableStyle() { Val = "PlainTable43" },
                    new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct }));

                var tableGrid = table.AppendChild(new TableGrid(new GridColumn()));
                //-----------------------------------------------------------------------------

                StyleDefinitionsPart part = wordDocument.MainDocumentPart.StyleDefinitionsPart;
                var styleid = "PlainTable43";
                var stylename = "Plain Table 43";

                // If the Styles part does not exist, add it and then add the style.
                if (part == null)
                {
                    part = AddStylesPartToPackage(wordDocument);
                    AddNewTableStyle(part, styleid, stylename);
                }
                else
                {
                    // If the style is not in the document, add it.
                    if (IsStyleIdInDocument(part, styleid) != true)
                    {
                        // No match on styleid, so let's try style name.
                        string styleidFromName = GetStyleIdFromStyleName(wordDocument, stylename);
                        if (styleidFromName == null)
                        {
                            AddNewTableStyle(part, styleid, stylename);
                        }
                        else
                            styleid = styleidFromName;
                    }
                }

                //-----------------------------------------------------------------------------

                var tr = new TableRow();
                var tc = new TableCell();

                var p = new Paragraph();
                var r = new Run();
                var text = new Text() { Text = "Working" };

                table.AppendChild(tr);
                tr.AppendChild(tc);
                tc.AppendChild(p);
                r.AppendChild(text);

                p.AppendChild(r);


                //----------------------------------------------------------------------------------

                tr = new TableRow();
                tc = new TableCell();

                p = new Paragraph();
                r = new Run();

                text = new Text() { Text = "Working better" };

                table.AppendChild(tr);
                tr.AppendChild(tc);
                tc.AppendChild(p);
                r.AppendChild(text);

                p.AppendChild(r);

                //--------------------------------------------------------------

                tr = new TableRow();
                tc = new TableCell();

                p = new Paragraph();
                r = new Run();

                text = new Text() { Text = "Working better" };

                table.AppendChild(tr);
                tr.AppendChild(tc);
                tc.AppendChild(p);
                r.AppendChild(text);

                p.AppendChild(r);


                tr = new TableRow();
                tc = new TableCell();

                p = new Paragraph();
                r = new Run();

                text = new Text() { Text = "Working better" };

                table.AppendChild(tr);
                tr.AppendChild(tc);
                tc.AppendChild(p);
                r.AppendChild(text);

                p.AppendChild(r);


                tr = new TableRow();
                tc = new TableCell();

                p = new Paragraph();
                r = new Run();

                text = new Text() { Text = "Working amazing sdgdsgdssssssssssssssssss" };

                table.AppendChild(tr);
                tr.AppendChild(tc);
                tc.AppendChild(p);
                r.AppendChild(text);

                p.AppendChild(r);
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
                var color = new Color() { Val = "FF7C80" };

                var run = new Run();
                var text = new Text();

                var highLight = new Highlight() { Val = HighlightColorValues.LightGray };

                var sectPr = new SectionProperties();

                var pgSz = new PageSize();
                var pgMar = new PageMargin();
                var cols = new Columns();
                var docGrid = new DocGrid();

                //----------------------------------------------------------------------------------------------

                body.AppendChild(p);

                p.AppendChild(run);
                run.AppendChild(rPr);
                rPr.AppendChild(color);
                rPr.AppendChild(highLight);

                run.AppendChild(text);

                text.Text = "Cmd";


            }
        }

        public static void blockQuote(string filepath)
        {
            using (WordprocessingDocument wordDocument =
WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {

                //creating the doc
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                Divs divs = new Divs();
                Div div = new Div();
                BlockQuote b = new BlockQuote() { Val = true };
                Paragraph p = new Paragraph();
                Run run = new Run(new Text() { Text = "hello!" });

                body.AppendChild(divs);
                divs.AppendChild(div);
                div.AppendChild(b);
                div.AppendChild(p);
                p.AppendChild(run);

                var p2 = new Paragraph();
                var run2 = new Run(new Text() { Text = "hello! this is worst" });

                body.AppendChild(p2);
                p2.AppendChild(run2);

            }
        }

       

    }

}
