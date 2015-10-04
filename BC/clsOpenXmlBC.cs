using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using AQ1Report.Utils;
//using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
//using E = DocumentFormat.OpenXml.OpenXmlElement;
//using A = DocumentFormat.OpenXml.OpenXmlAttribute;


namespace AQ1Report.BC
{
    class clsOpenXmlBC
    {
        //DOM approach
        public static void ReadExcelFileDOM(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
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
                Console.WriteLine();
                Console.ReadKey();
            }
        }

        //SAX approach
        public static void ReadExcelFileSAX(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        Console.Write(text + " ");
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }

        public static void ReadExcelFileStackTrace(string fileName)
        {
            //REFERENCE: http://stackoverflow.com/questions/23102010/open-xml-reading-from-excel-file

            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<Row>();

                    Console.WriteLine("Row count = {0}", rows.LongCount());
                    Console.WriteLine("Cell count = {0}", cells.LongCount());

                    // One way: go through each cell in the sheet
                    Console.WriteLine("Approach 1....");
                    foreach (Cell cell in cells)
                    {
                        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(cell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            Console.WriteLine("Shared string {0}: {1}", ssid, str);
                        }
                        else if (cell.CellValue != null)
                        {
                            Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
                        }
                    }

                    // Or... via each row
                    Console.WriteLine("");
                    Console.WriteLine("");
                    Console.WriteLine("Approach 2....");
                    foreach (Row row in rows)
                    {
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                Console.WriteLine("Shared string {0}: {1}", ssid, str);
                            }
                            else if (c.CellValue != null)
                            {
                                Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                            }
                        }
                    }
                }
            }
        }

        public static string GetCellValue(Cell c, SharedStringTable sst)
        {
            string cellValue = "~UNKNOWN~";

            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
            {
                int ssid = int.Parse(c.CellValue.Text);
                cellValue = sst.ChildElements[ssid].InnerText;
            }
            else if (c.CellValue != null)
            {
                cellValue = c.CellValue.Text;
            }

            return cellValue;
        }

        public static void ExcelBasics(string fileName)
        {
            //REFERENCE: https://msdn.microsoft.com/en-us/library/office/bb507946.aspx?cs-save-lang=1&cs-lang=csharp#code-snippet-6

            using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
            {
                Sheets sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;

                //Console.WriteLine("Sheet Name: Approach 1");
                //foreach (OpenXmlElement sheet in sheets)
                //{
                //    foreach (OpenXmlAttribute attr in sheet.GetAttributes())
                //    {
                //        Console.Write("{0}: {1} | ", attr.LocalName, attr.Value);
                //    }
                //    Console.WriteLine("");
                //}

                Console.WriteLine("");
                Console.WriteLine("Sheet Name: Approach 2");
                foreach (Sheet item in sheets)
                {
                    Console.WriteLine("ID: " + item.Id + " | " + item.Name);
                }
            }
        }

        public enum CellDataTypeEnum
        {
            //Boolean, Date, Error, InlineString, Number, SharedString, String
            //Number, SharedString
            AutoDetect, Number, SharedString
        }
        public static void SetCell(SharedStringTablePart shareStringPart, WorksheetPart worksheetPart, uint rowIndex, string columName, string value, CellDataTypeEnum dataType)
        {
            if (dataType == CellDataTypeEnum.AutoDetect)
            {
                if (clsValidation.IsNumeric(value)) dataType = CellDataTypeEnum.Number;
                else dataType = CellDataTypeEnum.SharedString;
            }
            // Insert the text into the SharedStringTablePart.
            //int index = clsOpenXmlBC.InsertSharedStringItem(value, shareStringPart);

            Cell cell;
            cell = InsertCellInWorksheet(worksheetPart, rowIndex, columName);
            //cell.CellValue = new CellValue(index.ToString());
            switch (dataType)
            {
                case CellDataTypeEnum.Number:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    cell.CellValue = new CellValue(value);
                    break;

                case CellDataTypeEnum.SharedString:
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    int index = clsOpenXmlBC.InsertSharedStringItem(value, shareStringPart);
                    cell.CellValue = new CellValue(index.ToString());
                    break;
                
                default:
                    throw new Exception("PROGRAMMING PENDING for this data type");
            }
        }

        //REFERENCE: https://msdn.microsoft.com/EN-US/library/office/cc861607.aspx
        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        //REFERENCE: https://msdn.microsoft.com/EN-US/library/office/cc861607.aspx
        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        //Defect 1: Fails for Columns AA and above.
        //Defect 1: Solution: http://stackoverflow.com/questions/14525573/open-xml-sdk-get-unreadable-content-error-when-trying-to-populate-more-than-2
        private static Cell InsertCellInWorksheet(WorksheetPart worksheetPart, uint rowIndex, string columnName)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    //Defect 1: START
                    //if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    if (GetExcelColumnNo(cell.CellReference.Value) > GetExcelColumnNo(cellReference))
                    //Defect 1: END
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        //REFERENCE: http://stackoverflow.com/questions/14525573/open-xml-sdk-get-unreadable-content-error-when-trying-to-populate-more-than-2
        //Renamed function name from ColumnNameParse to GetExcelColumnNo
        private static readonly string _Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static int GetExcelColumnNo(string excelColumnName)
        {
            // assumes value.Length is [1,3]
            // assumes value is uppercase
            var digits = excelColumnName.PadLeft(3).Select(x => _Alphabet.IndexOf(x));
            return digits.Aggregate(0, (current, index) => (current * 26) + (index + 1));
        }

        public static string GetExcelColumnName(string cellReference)
        {
            return cellReference.TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');
        }
    }
}