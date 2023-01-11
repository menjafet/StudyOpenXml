// See https://aka.ms/new-console-template for more information
using Documentxml;

using ClassLibrary1;

Console.WriteLine("Hello, World!");
//Documentxml.createDocument("/Users/fabianvalverde/Documents/StudyOpenXml/SampleFile.docx", "Style01", "Best Style");
//Documentxml.createCheckBox("C:/Users/Fabia/OneDrive/Documents/tests/SampleFile.docx","check",1,"Todosaliobien?");
//Documentxml.createDocument("C:/Users/Fabia/OneDrive/Documents/tests/createDocument.docx", "Heading1", "Normal Style");
//Documentxml.createTable("C:/Users/Fabia/OneDrive/Documents/tests/createTable.docx");
//DocManipulation.createCheckBox2("C:/Users/Fabia/OneDrive/Documents/tests/checkBox.docx");
//DocManipulation.changeBackgroundTable("C:/Users/Fabia/OneDrive/Documents/tests/BackgroundTable.docx");
//DocManipulation.highlightText("C:/Users/Fabia/OneDrive/Documents/tests/highLight.docx");
//DocManipulation.blockQuote(@"C:\\Users\\jaftb\\Documents\\StudyOpenXml\\tests\\BlockQuote.docx");
string document = @"C:\Users\jaftb\OneDrive\Escritorio\test\Word9.docx";
string document2 = @"C:\Users\jaftb\OneDrive\Escritorio\test\Word10.docx";
try
{
    File.Delete(document2);
    
}
finally
{
    File.Copy(document, document2);
}

string fileName = @"https://upload.wikimedia.org/wikipedia/commons/thumb/d/d7/Cristo_crucificado.jpg/240px-Cristo_crucificado.jpg";
ClassLibrary1.Class1.InsertAPicture(document2, fileName);
