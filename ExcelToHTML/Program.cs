using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
//using Microsoft.Office.Interop.Excel;

namespace ExcelToHTML
{
    class Program
    {
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }

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
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
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
        static void Main(string[] args)

        {
            /*var tableThree = new Aspose.Cells.Workbook("2019Table3.xlsx");
            //save XLS as HTML
            Aspose.Cells.HtmlSaveOptions htmlSaveOptions = new Aspose.Cells.HtmlSaveOptions();
            tableThree.Save("fascd", );
            //tableThree.Save("2019Table3(1).html");*/

            var doc = new HtmlDocument();

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("2019Table3.xlsx", false);

            spreadsheetDocument.ChangeDocumentType((SpreadsheetDocumentType)WordprocessingDocumentType.Document);


            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            workbookPart.Workbook = new Workbook();

            string html = "<head>";
            html += "<title>Page Title</title>";

            html += "</head><body>";
            html += "<p>";
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            //Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart).Worksheet;

            string dataText;

                String[] columnLetter = { "A", "B", "C", "D", "E", "F", "G", "I" }; //gotta be a better way then just putting the letter

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
            //IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
            html += "<table>";
            html += "Table 3. Incidence rates1 of nonfatal occupational injuries and illnesses by industry sector and employment size, California, 2019";
                for (int i = 3; i <= 38; i++) {
                    html += "<tr>";
                    foreach (string l in columnLetter)
                    {
                        html += "<td rowspan = 1 colspan = 0>";
                       
                        if (GetCellValue("2019Table3.xlsx", "SUM_3_06_2019", l + i.ToString()) != null) //so this doesnt work coolio and its 3AM so tmrw i guess
                        {
                     
                            dataText = GetCellValue("2019Table3.xlsx", "SUM_3_06_2019", l + i.ToString());
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
