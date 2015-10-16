using AQ1Report.Entities;
using AQ1Report.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;


namespace AQ1Report.BC
{
    class clsAQ1ReportBC
    {
        #region "Enums"
        private enum RowTypeEnum
        {
            UNKNOWN_ROW_TYPE,
            HeaderRow,
            DetailRow,
            FooterRow,
        }

        private enum RowSubTypeEnum
        {
            UNKNOWN_ROW_SUB_TYPE,
            HeaderRow1, HeaderRow2, HeaderRow3, HeaderRow4, HeaderRow5, HeaderRow6, HeaderRow7, HeaderRow8,
            DetailRow1, DetailRow2,
            FooterRow1, FooterRow2, FooterRow3
        }

        private enum TypeEnum
        {
            UNKNOWN_TYPE,
            METER_READING_NOT_ENTERED,
            PERMANENTLY_CLOSED
        }
        #endregion

        public static void ProcessAQ1Excel(string inputFileName, string outputFileName)
        {
            List<clsAQ1ReportDetailRow> AQ1ReportDetailRows;

            AQ1ReportDetailRows = ReadAQ1Excel(inputFileName);

            CreateAQ1Report(outputFileName, AQ1ReportDetailRows);
        }


        #region "AQ1: Read from Excel"
        private static List<clsAQ1ReportDetailRow> ReadAQ1Excel(string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    WorkbookPart wbp = doc.WorkbookPart;
                    Sheets sheets = wbp.Workbook.Sheets;
                    WorksheetPart wsp;

                    Worksheet ws;
                    List<clsAQ1ReportDetailRow> detailRows = new List<clsAQ1ReportDetailRow>();
                    int sheetNo = 0;

                    Console.WriteLine("");
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Reading input file....");
                    Console.ResetColor();
                    foreach (Sheet item in sheets)
                    {
                        sheetNo++;

                        wsp = (WorksheetPart)wbp.GetPartById(item.Id);
                        ws = wsp.Worksheet;

                        Console.Write("Reading sheet no " + sheetNo.ToString() + "....");
                        ReadAQ1ExcelWorkSheet(ws, sheetNo, sst, detailRows);
                        Console.ForegroundColor = ConsoleColor.DarkGreen;
                        Console.WriteLine("DONE!");
                        Console.ResetColor();
                    }

                    return detailRows;
                }
            }
        }

        private static void ReadAQ1ExcelWorkSheet(Worksheet sheet, int sheetNo, SharedStringTable sst, List<clsAQ1ReportDetailRow> detailRows)
        {
            Debug.WriteLine("ReadAQ1ExcelWorkSheet");

            var cells = sheet.Descendants<Cell>();
            var rows = sheet.Descendants<Row>();

            Debug.WriteLine("Row Count = {0}", rows.LongCount());
            Debug.WriteLine("Cell Count = {0}", cells.LongCount());
            Debug.WriteLine("");

            clsAQ1ReportHeaderRow headerRow = new clsAQ1ReportHeaderRow(sheetNo);
            clsAQ1ReportDetailRow detailRow = null;
            int rowCount = rows.Count();

            RowTypeEnum rowType = RowTypeEnum.UNKNOWN_ROW_TYPE;
            RowSubTypeEnum rowSubType = RowSubTypeEnum.UNKNOWN_ROW_SUB_TYPE;

            foreach (Row row in rows)
            {
                SetRowTypeSubType(ref rowType, ref rowSubType, row, rowCount, sst);
                Debug.WriteLine("\t" + "rowType: " + rowType.ToString() + ", rowSubType: " + rowSubType.ToString());

                switch (rowType)
                {
                    case RowTypeEnum.HeaderRow:
                        PopulateAQ1HeaderRow(sst, headerRow, rowSubType, row);
                        continue;

                    case RowTypeEnum.DetailRow:
                        if (rowSubType == RowSubTypeEnum.DetailRow1)
                        {
                            detailRow = new clsAQ1ReportDetailRow(headerRow, row.RowIndex);
                            detailRows.Add(detailRow);
                        }

                        PopulateAQ1DetailRow(sst, detailRow, rowSubType, row);

                        break;

                    case RowTypeEnum.FooterRow:
                        continue;

                    default:
                        throw new Exception("UNKNOWN Row Type");
                }
            }
        }

        private static void SetRowTypeSubType(ref RowTypeEnum rowType, ref RowSubTypeEnum rowSubType, Row row, int rowCount, SharedStringTable sst)
        {
            Debug.WriteLine("SetRowTypeSubType");

            rowType = RowTypeEnum.UNKNOWN_ROW_TYPE;
            rowSubType = RowSubTypeEnum.UNKNOWN_ROW_SUB_TYPE;

            if (row.RowIndex == 1)
            {
                rowType = RowTypeEnum.HeaderRow;
                rowSubType = RowSubTypeEnum.HeaderRow1;
            }
            else if (row.RowIndex == 2)
            {
                rowType = RowTypeEnum.HeaderRow;
                rowSubType = RowSubTypeEnum.HeaderRow2;
            }
            else if (row.RowIndex == 3)
            {
                rowType = RowTypeEnum.HeaderRow;
                rowSubType = RowSubTypeEnum.HeaderRow3;
            }
            else if (row.RowIndex == 4)
            {
                rowType = RowTypeEnum.HeaderRow;
                rowSubType = RowSubTypeEnum.HeaderRow4;
            }
            else if (row.RowIndex == 5)
            {
                rowType = RowTypeEnum.HeaderRow;
                rowSubType = RowSubTypeEnum.HeaderRow5;
            }
            else if (row.RowIndex == 6)
            {
                rowType = RowTypeEnum.HeaderRow;
                rowSubType = RowSubTypeEnum.HeaderRow6;
            }
            else if (row.RowIndex == 7)
            {
                rowType = RowTypeEnum.HeaderRow;
                rowSubType = RowSubTypeEnum.HeaderRow7;
            }
            else if (row.RowIndex == 8)
            {
                rowType = RowTypeEnum.HeaderRow;
                rowSubType = RowSubTypeEnum.HeaderRow8;
            }
            else if (row.RowIndex == rowCount)
            {
                rowType = RowTypeEnum.FooterRow;
                rowSubType = RowSubTypeEnum.FooterRow3;
            }
            else
            {
                Cell cell = row.Elements<Cell>().ElementAt(0);
                string cellColumn = clsOpenXmlBC.GetExcelColumnName(cell.CellReference.ToString());
                string cellValue = clsOpenXmlBC.GetCellValue(cell, sst);
                Debug.WriteLine("\t" + "cell.CellReference" + cell.CellReference.ToString() + ", cellColumn: " + cellColumn + ", cellValue: " + cellValue);
                if ((cellColumn == "A") && (clsValidation.IsInt32(cellValue)))
                {
                    rowType = RowTypeEnum.DetailRow;
                    rowSubType = RowSubTypeEnum.DetailRow1;
                }
                else if ((cellColumn == "B") && (cellValue == "Total Meter Readings Entered :"))
                {
                    rowType = RowTypeEnum.FooterRow;
                    rowSubType = RowSubTypeEnum.FooterRow1;
                }
                else if (((cellColumn == "B") && (cellValue == "Total Meter Readings Not Entered"))
                    || ((cellColumn == "A") && (cellValue == "Total Meter Readings Not Entered")))
                {
                    rowType = RowTypeEnum.FooterRow;
                    rowSubType = RowSubTypeEnum.FooterRow2;
                }
                else
                {
                    rowType = RowTypeEnum.DetailRow;
                    rowSubType = RowSubTypeEnum.DetailRow2;
                }
            }
        }

        private static void PopulateAQ1HeaderRow(SharedStringTable sst, clsAQ1ReportHeaderRow objAQ1HeaderRow, RowSubTypeEnum rowSubType, Row row)
        {
            Debug.WriteLine("PopulateAQ1HeaderRow");

            if (rowSubType == RowSubTypeEnum.HeaderRow1)
            {
                objAQ1HeaderRow.Date = clsOpenXmlBC.GetCellValue(row.Elements<Cell>().ElementAt(2), sst).TrimStart(' ', ':');
            }
            else if (rowSubType == RowSubTypeEnum.HeaderRow2)
            {
                objAQ1HeaderRow.PageNo = clsOpenXmlBC.GetCellValue(row.Elements<Cell>().ElementAt(2), sst);
            }
            else if (rowSubType == RowSubTypeEnum.HeaderRow3)
            {
                //Empty Row. No action;
            }
            else if (rowSubType == RowSubTypeEnum.HeaderRow4)
            {
                objAQ1HeaderRow.GeneratedBy = clsOpenXmlBC.GetCellValue(row.Elements<Cell>().ElementAt(4), sst);
            }
            else if (rowSubType == RowSubTypeEnum.HeaderRow5)
            {
                objAQ1HeaderRow.WardName = clsOpenXmlBC.GetCellValue(row.Elements<Cell>().ElementAt(1), sst).TrimStart(' ', ':');
            }
            else if (rowSubType == RowSubTypeEnum.HeaderRow6)
            {
                objAQ1HeaderRow.MeterBinderNo = clsOpenXmlBC.GetCellValue(row.Elements<Cell>().ElementAt(1), sst).TrimStart(' ', ':');
                objAQ1HeaderRow.ReadingCycle = clsOpenXmlBC.GetCellValue(row.Elements<Cell>().ElementAt(4), sst);
                objAQ1HeaderRow.Year = clsOpenXmlBC.GetCellValue(row.Elements<Cell>().ElementAt(7), sst);
            }
        }

        private static void PopulateAQ1DetailRow(SharedStringTable sst, clsAQ1ReportDetailRow objAQ1ReportRow, RowSubTypeEnum rowSubType, Row row)
        {
            Debug.WriteLine("PopulateAQ1DetailRow");
            string cellValue;
            string column;

            foreach (Cell c in row.Elements<Cell>())
            {
                cellValue = clsOpenXmlBC.GetCellValue(c, sst);

                column = clsOpenXmlBC.GetExcelColumnName(c.CellReference.ToString());
                Debug.WriteLine("\t" + "CellReference: " + c.CellReference.ToString() + ", column: " + column);

                switch (column)
                {
                    case "A": objAQ1ReportRow.A_row1_Folio = cellValue; break;
                    case "B": objAQ1ReportRow.B_row1_CCN = cellValue; break;
                    case "C": objAQ1ReportRow.C_row1_CCNStat = cellValue; break;
                    case "D": objAQ1ReportRow.D_row1_CCNLink = cellValue; break;
                    case "E": objAQ1ReportRow.E_row1_CCN10 = cellValue; break;
                    case "F": objAQ1ReportRow.F_row1_MtrStat = cellValue; break;
                    case "G": objAQ1ReportRow.G_row1_GAP = cellValue; break;
                    case "H":
                        if (rowSubType == RowSubTypeEnum.DetailRow1)
                            objAQ1ReportRow.H_row1_CurrentDate = cellValue;
                        else if (rowSubType == RowSubTypeEnum.DetailRow2)
                            objAQ1ReportRow.H_row2_CurrentRdg = cellValue;

                        break;
                    case "I":
                        if (rowSubType == RowSubTypeEnum.DetailRow1)
                            objAQ1ReportRow.I_row1_PreviousDate = cellValue;
                        else if (rowSubType == RowSubTypeEnum.DetailRow2)
                            objAQ1ReportRow.I_row2_PreviousRdg = cellValue;

                        break;
                    case "J": objAQ1ReportRow.J_row1_CutRemvDt = cellValue; break;
                    case "K": objAQ1ReportRow.K_row1_RstrRplcDt = cellValue; break;
                    case "L": objAQ1ReportRow.L_row1_CsmpByMeter = cellValue; break;
                    case "M": objAQ1ReportRow.M_row1_Days = cellValue; break;
                    case "N": objAQ1ReportRow.N_row1_CsmpBilled = cellValue; break;
                    case "O": objAQ1ReportRow.O_row1_Water = cellValue; break;
                    case "P": objAQ1ReportRow.P_row1_Sewarage = cellValue; break;
                    case "Q": objAQ1ReportRow.Q_row1_Rent = cellValue; break;
                    case "R": objAQ1ReportRow.R_row1_BillAmt = cellValue; break;
                    case "S": objAQ1ReportRow.S_row1_Additional = cellValue; break;
                    case "T": objAQ1ReportRow.T_row1_Flat = cellValue; break;
                    case "U": objAQ1ReportRow.U_row1_Flag = cellValue; break;
                    case "V": objAQ1ReportRow.V_row1_GroupCode = cellValue; break;
                    case "W": objAQ1ReportRow.W_row1_RateCharge = cellValue; break;
                    case "X":
                        if (rowSubType == RowSubTypeEnum.DetailRow1)
                            objAQ1ReportRow.X_row1_CutDate = cellValue;
                        else if (rowSubType == RowSubTypeEnum.DetailRow2)
                            objAQ1ReportRow.X_row2_Reason_Part2 = cellValue;

                        break;
                    case "Y": objAQ1ReportRow.Y_row1_Reason = cellValue; break;
                }
            }
        }
        #endregion


        #region "AQ1: Write to Excel"
        private static void CreateAQ1Report(string fileName, List<clsAQ1ReportDetailRow> AQ1ReportDetailRows)
        {
            //REFERENCE: https://msdn.microsoft.com/en-us/library/office/ff478153.aspx

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "AQ1-Consolidated"
            };
            sheets.Append(sheet);

            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            Console.WriteLine("");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Creating output file....");
            Console.ResetColor();
            uint rowIndex = 1;

            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "A", "Meter Binder No", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "B", "Folio", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "C", "CCN", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "D", "CCN Link", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "E", "CCN Stat", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "F", "Group Code", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "G", "Mtr Stat", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "H", "GAP", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "I", "Type", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "J", "Cut Date", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "K", "Cutoff Reason", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "L", "Current Date", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "M", "Current Reading", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "N", "Previous Date", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "O", "Previous Rdg", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "P", "Rate Charge", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "Q", "Cut Remv Dt", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "R", "Cut Remv", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "S", "Rstr Rplc Dt", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "T", "Rstr Rplc", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "U", "Csmp By Meter", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "V", "Days", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "W", "Csmp Billed", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "X", "Water", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "Y", "Sewarage", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "Z", "Rent", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AA", "Bill Amt", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AB", "Additional", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AC", "Flat", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AD", "Date", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AE", "Page No", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AF", "Generated By", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AG", "Ward Name", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AH", "Reading Cycle", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AI", "Year", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AJ", "Sheet No", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AK", "Row No", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AL", "CCN 10", clsOpenXmlBC.CellDataTypeEnum.SharedString);
            clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AM", "Flag", clsOpenXmlBC.CellDataTypeEnum.SharedString);

            TypeEnum type = TypeEnum.UNKNOWN_TYPE;
            foreach (clsAQ1ReportDetailRow detailRow in AQ1ReportDetailRows)
            {
                Console.Write("Writing row no " + rowIndex.ToString() + " of " + AQ1ReportDetailRows.Count.ToString() + "....");
                rowIndex++;

                type = TypeEnum.UNKNOWN_TYPE;
                if ((detailRow.G_row1_GAP == "Meter") && (detailRow.H_row1_CurrentDate == "Reading") && (detailRow.I_row1_PreviousDate == "Not Entered"))
                    type = TypeEnum.METER_READING_NOT_ENTERED;
                else if ((detailRow.H_row1_CurrentDate == "Permanently") && (detailRow.I_row1_PreviousDate == "Closed"))
                    type = TypeEnum.PERMANENTLY_CLOSED;


                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "A", detailRow.HeaderRow.MeterBinderNo.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "B", detailRow.A_row1_Folio.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "C", detailRow.B_row1_CCN.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "D", detailRow.D_row1_CCNLink.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "E", detailRow.C_row1_CCNStat.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "F", detailRow.V_row1_GroupCode.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "G", detailRow.F_row1_MtrStat.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                
                if (type != TypeEnum.METER_READING_NOT_ENTERED)
                    clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "H", detailRow.G_row1_GAP.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);

                if (type == TypeEnum.METER_READING_NOT_ENTERED)
                    clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "I", "Meter Reading Not Entered", clsOpenXmlBC.CellDataTypeEnum.SharedString);
                else if (type == TypeEnum.PERMANENTLY_CLOSED)
                    clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "I", "Permanently Cutoff", clsOpenXmlBC.CellDataTypeEnum.SharedString); clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "K", detailRow.X_row1_CutDate.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);

                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "J", detailRow.X_row1_CutDate.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "K", detailRow.Y_row1_Reason.ToString() + " " + detailRow.X_row2_Reason_Part2.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);

                if ((type != TypeEnum.METER_READING_NOT_ENTERED) && (type != TypeEnum.PERMANENTLY_CLOSED))
                    clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "L", detailRow.H_row1_CurrentDate.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);

                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "M", detailRow.H_row2_CurrentRdg.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                
                if ((type != TypeEnum.METER_READING_NOT_ENTERED) && (type != TypeEnum.PERMANENTLY_CLOSED))
                    clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "N", detailRow.I_row1_PreviousDate.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);

                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "O", detailRow.I_row2_PreviousRdg.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "P", detailRow.W_row1_RateCharge.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "Q", detailRow.J_row1_CutRemvDt.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "R", detailRow.J_row2_CutRemv.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "S", detailRow.K_row1_RstrRplcDt.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "T", detailRow.K_row2_RstrRplc.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "U", detailRow.L_row1_CsmpByMeter.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "V", detailRow.M_row1_Days.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "W", detailRow.N_row1_CsmpBilled.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "X", detailRow.O_row1_Water.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "Y", detailRow.P_row1_Sewarage.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "Z", detailRow.Q_row1_Rent.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AA", detailRow.R_row1_BillAmt.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AB", detailRow.S_row1_Additional.ToString(), clsOpenXmlBC.CellDataTypeEnum.AutoDetect);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AC", detailRow.T_row1_Flat.ToString(), clsOpenXmlBC.CellDataTypeEnum.AutoDetect);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AD", detailRow.HeaderRow.Date.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AE", detailRow.HeaderRow.PageNo.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AF", detailRow.HeaderRow.GeneratedBy.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AG", detailRow.HeaderRow.WardName.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AH", detailRow.HeaderRow.ReadingCycle.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AI", detailRow.HeaderRow.Year.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AJ", detailRow.HeaderRow.ExcelWorkSheetNo.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AK", detailRow.ExcelRowNo.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AL", detailRow.E_row1_CCN10.ToString(), clsOpenXmlBC.CellDataTypeEnum.Number);
                clsOpenXmlBC.SetCell(shareStringPart, worksheetPart, rowIndex, "AM", detailRow.U_row1_Flag.ToString(), clsOpenXmlBC.CellDataTypeEnum.SharedString);

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("DONE!");
                Console.ResetColor();
            }
            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        }
        #endregion
    }
}
