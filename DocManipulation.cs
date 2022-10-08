using System;
using System.Security.AccessControl;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Security.Cryptography.X509Certificates;
using DocumentFormat.OpenXml.Vml;

namespace StudyOpenXml
{
    public class DocManipulation
    {
        public DocManipulation()
        {
        }


        public static void CreateDocument(string filepath, string styleid, string stylename)
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

                ////This could be the way to loop and create an .md from .docx

                //Paragraph[] paraArray = new Paragraph[2];
                //Run[] runArray = new Run[2];
                //paraArray[0] = body.AppendChild(new Paragraph());
                //paraArray[1] = body.AppendChild(new Paragraph());
                //runArray[0] = paraArray[0].AppendChild(new Run());
                //runArray[1] = paraArray[1].AppendChild(new Run());

                //runArray[0].AppendChild(new Text("Everything Working"));
                //runArray[1].AppendChild(new Text("We can work now"));

                ////Lets format a Paragraph
                //ParagraphProperties[] pPr = new ParagraphProperties[2];
                ParagraphProperties paraProperties = new ParagraphProperties();



                //Get the first Paragraph of the document
                Paragraph p = wordDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>()
                .ElementAtOrDefault(0);

                //Every element with the ParagraphProperties type
                if(p.Elements<ParagraphProperties>().Count() == 0)
                {
                    //Add ParagraphProperties to the first Paragraph
                    p.PrependChild<ParagraphProperties>(new ParagraphProperties());
                }

                //Getting the first element Paragraph Properties
                paraProperties = p.Elements<ParagraphProperties>().First();

                //Looking for the stylespart (not exactly the styles itself) of the document
                //(This is a new document, styles part aren't by default)
                StyleDefinitionsPart part = wordDocument.MainDocumentPart.StyleDefinitionsPart;

                if(part == null)
                {
                    part = AddStylesPartToPackage(wordDocument);
                    AddNewStyle(part, styleid, stylename);
                }
                else
                {
                    if (IsStyleIdInDocument(part, styleid) != true)
                    {
                        // No match on styleid, so let's try style name.
                        string styleidFromName = GetStyleIdFromStyleName(wordDocument, stylename);
                        if (styleidFromName == null)
                        {
                            AddNewStyle(part, styleid, stylename);
                        }
                        else
                            styleid = styleidFromName;
                    }
                }
                // Set the style of the paragraph.
                paraProperties.ParagraphStyleId = new ParagraphStyleId { Val = styleid };
            }
        }
        //inside method
      /*  string filename = @"C:\Users\Public\Documents\ApplyStyleToParagraph.docx";

    using (WordprocessingDocument doc =
        WordprocessingDocument.Open(filename, true))
    {
        // Get the first paragraph.
        Paragraph p =
          doc.MainDocumentPart.Document.Body.Descendants<Paragraph>()
          .ElementAtOrDefault(1);

        // Check for a null reference. 
        if (p == null)
        {
            throw new ArgumentOutOfRangeException("p",
                "Paragraph was not found.");
}

ApplyStyleToParagraph(doc, "OverdueAmount", "Overdue Amount", p);
    }*/

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
    public static void CreateTable(string filepath)
    {
            // Use the file name and path passed in as an argument 
            // to open an existing Word 2007 document.

            using (WordprocessingDocument wordDocument =
                        WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                //creating the doc
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                Table table = new Table();

                //Table borders
                var borderColor = "FF8000";
                var tblBorder = new TableBorders();


                var topBorder = new TopBorder();
                topBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                topBorder.Size = 8;
                topBorder.Color = borderColor;

                var bottomBorder = new BottomBorder();
                bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                bottomBorder.Size = 8;
                bottomBorder.Color = borderColor;

                var rightBorder = new RightBorder();
                rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                rightBorder.Size = 8;
                rightBorder.Color = borderColor;

                var leftBorder = new LeftBorder();
                leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                leftBorder.Size = 8;
                leftBorder.Color = borderColor;

                var insideHorizontalBorder = new InsideHorizontalBorder();
                insideHorizontalBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                insideHorizontalBorder.Size = 8;
                insideHorizontalBorder.Color = borderColor;

                var insideVerticalBorder = new InsideVerticalBorder();
                insideVerticalBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                insideVerticalBorder.Size = 8;
                insideVerticalBorder.Color = borderColor;


                tblBorder.AppendChild(topBorder);
                tblBorder.AppendChild(bottomBorder);
                tblBorder.AppendChild(rightBorder);
                tblBorder.AppendChild(leftBorder);
                tblBorder.AppendChild(insideHorizontalBorder);
                tblBorder.AppendChild(insideVerticalBorder);


                //-----------------------------------------------------------------------------------------
                //Below is what i need

                var tblProp = new TableProperties();
                var tableGrid = new TableGrid();
                var gridCol1 = new GridColumn() { Width = "1" };
                var gridCol2 = new GridColumn() { Width = "1" };

                TableLayout tl = new TableLayout() { Type = TableLayoutValues.Fixed };
                tblProp.TableLayout = tl;

                tableGrid.AppendChild(gridCol1);
                tableGrid.AppendChild(gridCol2);

                //same here
                tblProp.AppendChild(tblBorder);


                table.Append(tableGrid);
                

                table.AppendChild(tblProp);

                var row1 = new TableRow();

                var cell1 = new TableCell();


                var cellProp = new TableCellProperties();
                var cellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Auto, Width = "900" };

                var para = new Paragraph(new Run(new Text("edited2xcgsdgdsgdsgdsgfdsgdfghfdgdsgsdhfsd")));

                cell1.Append(para);

                row1.Append(cell1);

                TableCell cell2 = new TableCell(cell1.OuterXml);

                cellProp.AppendChild(cellWidth);

                cell1.AppendChild(cellProp);

                row1.Append(cell2);

                table.Append(row1);

                body.Append(table);
        }
    }
        //Want to try this on windows machine
        public static void createTable2(String filepath)
        {

            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {

                //This piece is adding the main part of the document
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                //Then, this is creating the structure to manipulate every part of the document
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                Table tbl = new Table();

                TableProperties tableProp = new TableProperties();
                TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };

                // Make the table width 100% of the page width.
                TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

                tableProp.Append(tableStyle, tableWidth);

                tbl.AppendChild(tableProp);

                //Add n columns to table
                TableGrid tg = new TableGrid(new GridColumn(), new GridColumn());

                tbl.AppendChild(tg);

                TableRow tr1 = new TableRow();

                //I Manually adjust width of the first column
                TableCell tc1 = new TableCell(new TableCellWidth { Type = TableWidthUnitValues.Auto }, new Paragraph(new Run(new Text("sdgdsgsdgsdgfds"))));

                //All other column are adjusted based on their content
                TableCell tc2 = new TableCell(new TableCellWidth { Type = TableWidthUnitValues.Auto}, new Paragraph(new Run(new Text("Title"))));

                tr1.Append(tc1, tc2);
                tbl.AppendChild(tr1);
                body.Append(tbl);
            }

            //This method is only used for headers, while regular rows cells contain no TableCellProperties

        }

        //public static TableCellProperties GenerateTableCellPropsWithWidth(string width = null)
        //{
        //    TableCellProperties tcp = new TableCellProperties();
        //    tcp.AppendChild(width.IsNullOrEmpty()
        //        ? new TableCellWidth { Type = TableWidthUnitValues.Auto }
        //        : new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = width });
        //    return tcp;
        //}
    }
}
