using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using AQ1Report.Utils;


namespace AQ1Report.BC
{
    class clsOpenXmlBC
    {
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

            Cell cell;
            cell = InsertCellInWorksheet(worksheetPart, rowIndex, columName);

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