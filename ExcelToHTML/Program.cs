using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Cells;

//using Microsoft.Office.Interop.Excel;

namespace ExcelToHTML
{
    class Program
    {
        public static string GetCellValue(string fileName,
        string sheetName,
        string addressName)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart =
                    (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                DocumentFormat.OpenXml.Spreadsheet.Cell theCell = wsPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell does not exist, return an empty string.
                if (theCell != null)
                {
                    value = theCell.InnerText;

                    // If the cell represents an integer number, you are done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and 
                    // Booleans individually. For shared strings, the code 
                    // looks up the corresponding value in the shared string 
                    // table. For Booleans, the code converts the value into 
                    // the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For shared strings, look up the value in the
                                // shared strings table.
                                var stringTable =
                                    wbPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                // If the shared string table is missing, something 
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in 
                                // the table.
                                if (stringTable != null)
                                {
                                    value =
                                        stringTable.SharedStringTable
                                        .ElementAt(int.Parse(value)).InnerText;
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
                document.Close();
            }
            return value;
        }

        public static string fileName = "2019Table3.xlsx";
        static void Main(string[] args)

        {
            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(fileName);

            //Access the first worksheet
            Aspose.Cells.Worksheet ws = wb.Worksheets[0];

            //Access last cell inside the worksheet
            Aspose.Cells.Cell cell = ws.Cells.LastCell;
            //Number of rows inside the worksheet
            int rowCount = cell.Row + 1;

            //Number of columns inside the worksheet
            Console.WriteLine("Number of Columns: " + (cell.Column + 1));
            /*var tableThree = new Aspose.Cells.Workbook("2019Table3.xlsx");
            //save XLS as HTML
            Aspose.Cells.HtmlSaveOptions htmlSaveOptions = new Aspose.Cells.HtmlSaveOptions();
            tableThree.Save("fascd", );
            //tableThree.Save("2019Table3(1).html");*/

            var doc = new HtmlDocument();

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false);

            spreadsheetDocument.ChangeDocumentType((SpreadsheetDocumentType)WordprocessingDocumentType.Document);


            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            string html = "<head>";
            html += "<title>Page Title</title>";

            html += "</head><body>";
            html += "<p>";
            //WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            //SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            string dataText;
            int colCount = 0;

            string[] columnLetter = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}; //gotta be a better way then just putting the letter
            //then access to the worksheet part
            IEnumerable<WorksheetPart> worksheetPart = workbookPart.WorksheetParts;

            /*foreach (WorksheetPart WSP in worksheetPart)
            {
                //find sheet data
                IEnumerable<SheetData> sheetData = WSP.Worksheet.Elements<SheetData>();
                // Iterate through every sheet inside Excel sheet
                foreach (SheetData SD in sheetData)
                {
                    IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> row = SD.Elements<Row>(); // Get the row IEnumerator
                    rowCount = row.Count(); ; // Will give you the count of rows
                }
            }*/

            int columnCount = 0;

            string[] columnsInDoc = new string[columnCount];
            for (int i = 0; i < columnCount; i++)
            {
                columnLetter[i] = columnsInDoc[i];
            }
            Console.WriteLine(columnCount);

            foreach (string t in columnsInDoc)
            {
                Console.WriteLine(t.ToString());
            }

            /*foreach (Row r in sheetData.Elements<Row>())
            {
                foreach (Cell c in r.Elements<Cell>())
                {
                    foreach (string l in columnLetter)
                    {
                        for (int i = 1; i <= 38; i++) //there's gotta be a more efficient way then to just go count the number of rows why am I so bad at coding 
                        {
                            if (c.DataType != null)
                            {
                                dataText = GetCellValue("2019Table3.xlsx", "SUM_3_06_2019", l + i.ToString());
                                html += "   ";
                                html += dataText;
                            }
                            else
                            {
                                html += "\r\n";
                            }
                        }

                    }
                }
            }*/

            html += "<table>";
            
                for (int i = 1; i <= rowCount; i++) {
                    html += "<tr>";
                    foreach (string l in columnLetter)
                    {
                        html += "<td rowspan = 1 colspan = 0>";
                       
                        if (GetCellValue(fileName, "SUM_3_06_2019", l + i.ToString()) != null) 
                        {
                            dataText = GetCellValue(fileName, "SUM_3_06_2019", l + i.ToString());
                            html += " ";
                            html += dataText;
                            html +=  "</td>";
                        }
                        else
                        { 
                            html += "<br>";
                            html += "</td>";
                        }

                    }
                    html += "</tr>";
                }
                html += "</table>";
                html += "<p>";
                html += "</body>";

                doc.LoadHtml(html);
                using (FileStream fs = new FileStream("afas.html", FileMode.Create))
                using (StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8) { AutoFlush = true })
                {
                    doc.Save(sw);
                }
            }
        }
    }
