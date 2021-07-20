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
            var tableThree = new Aspose.Cells.Workbook("2019Table3.xlsx");
            //save XLS as HTML
            tableThree.Save("2019Table3.html", Aspose.Cells.SaveFormat.Auto);
            /*SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("2019Table3.xlsx", false);
            spreadsheetDocument.ChangeDocumentType((SpreadsheetDocumentType)WordprocessingDocumentType.Document);
            string html =
        @"<!DOCTYPE html>
<html>
<body>
	<h1>This is <b>bold</b> heading</h1>
	<p>" + spreadsheetDocument + "paragraph</p>" +
	"<h2>This is <i>italic</i> heading</h2>" +
	"<h2>This is new heading</h2>" +
"</body>" +
"</html> ";

            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            workbookPart.Workbook = new Workbook();
            var doc = new HtmlDocument();

            doc.Load(html);
            doc.Save("2019Table3.html"); 

            var newDoc = new HtmlDocument();
            newDoc.DetectEncodingAndLoad("2019Table3.html");
            var node = newDoc.DocumentNode.SelectSingleNode("//body");*/





            /*WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            string text;
            foreach (Row r in sheetData.Elements<Row>())
            {
                foreach (Cell c in r.Elements<Cell>())
                {
                    text = c.CellValue.Text;
                    Console.Write(text + " ");
                }
            }

            FileStream sw = new FileStream("2019Table3.html", FileMode.Create);
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
        /*        static void ReadExcelFile()
                {
                    try
                    {
                        //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                        using (SpreadsheetDocument doc = SpreadsheetDocument.Open("2019Table3.xls", false))
                        {
                            //create the object for workbook part  
                            WorkbookPart workbookPart = doc.WorkbookPart;
                            Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                            StringBuilder excelResult = new StringBuilder();

                            //using for each loop to get the sheet from the sheetcollection  
                            foreach (Sheet thesheet in thesheetcollection)
                            {
                                excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                                excelResult.AppendLine("----------------------------------------------- ");
                                //statement to get the worksheet object by using the sheet id  
                                Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                                SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                                foreach (Row thecurrentrow in thesheetdata)
                                {
                                    foreach (Cell thecurrentcell in thecurrentrow)
                                    {
                                        //statement to take the integer value  
                                        string currentcellvalue = string.Empty;
                                        if (thecurrentcell.DataType != null)
                                        {
                                            if (thecurrentcell.DataType == CellValues.SharedString)
                                            {
                                                int id;
                                                if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                                {
                                                    SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                                    if (item.Text != null)
                                                    {
                                                        //code to take the string value  
                                                        excelResult.Append(item.Text.Text + " ");
                                                    }
                                                    else if (item.InnerText != null)
                                                    {
                                                        currentcellvalue = item.InnerText;
                                                    }
                                                    else if (item.InnerXml != null)
                                                    {
                                                        currentcellvalue = item.InnerXml;
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + " ");
                                        }
                                    }
                                    excelResult.AppendLine();
                                }
                                excelResult.Append("");
                                Console.WriteLine(excelResult.ToString());
                                Console.ReadLine();
                            }
                        }
                    }
                    catch (Exception)
                    {

                    }
                }*/
    }
}
