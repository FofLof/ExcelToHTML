using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using System.IO;
using System.Linq;
//using Microsoft.Office.Interop.Excel;

namespace ExcelToHTML
{
    class Program
    {
        public static string filePath = @"C:\Users\Ethan\source\repos\ExcelToHTML\ExcelToHTML\2019Table3.xls";
        static void Main(string[] args)

        {
            /*var tableThree = new Aspose.Cells.Workbook("2019Table3.xlsx");
            //save XLS as HTML
            Aspose.Cells.HtmlSaveOptions htmlSaveOptions = new Aspose.Cells.HtmlSaveOptions();
            tableThree.Save("fascd", );
            //tableThree.Save("2019Table3(1).html");*/




            var doc = new HtmlDocument();
            /*doc.Load(path);
            doc.Save("test.html");*/

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("2019Table3.xlsx", false);

            spreadsheetDocument.ChangeDocumentType((SpreadsheetDocumentType)WordprocessingDocumentType.Document);


                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                workbookPart.Workbook = new Workbook();
            /*var doc = new HtmlDocument();

            doc.Load(html);
            doc.Save("2019Table3.html");

            var newDoc = new HtmlDocument();
            newDoc.DetectEncodingAndLoad("2019Table3.html");
            var node = newDoc.DocumentNode.SelectSingleNode("//body");
*/


            string html = "<head>";
            html += "<title>Page Title</title>";

            html += "</head><body>";
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            string dataText;
            foreach (Row r in sheetData.Elements<Row>())
              {
                 foreach (Cell c in r.Elements<Cell>())
                 {
                    dataText = c.CellValue.Text;
                    html += dataText;
                 }
             }
              
                html += "</body>";

                doc.LoadHtml(html);
                using (FileStream fs = new FileStream("afsc.html", FileMode.Create))
                using (StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8) { AutoFlush = true })
                {
                    doc.Save(sw);

                /*FileStream sw = new FileStream("2019Table3.html", FileMode.Create);
                //htmlDoc.Save(sw);
                sw.Close();


                string html = "<head>";
                string style = "text-align:center";

                html += "<title>Page Title</title>";
                html += "<style>" + style + "</style>";
                html += "</head><body>";
                html +=
                html += "</body>";


                var doc = new HtmlDocument();
                doc.Load("2019Table2.xls");
*/


            }
        }
    }
}
