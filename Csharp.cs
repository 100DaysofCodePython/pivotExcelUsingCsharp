// Requires reference: Microsoft.Office.Interop.Excel
// Target .NET Framework (4.7.2 / 4.8). Run on Windows with Excel installed (x64 recommended).

using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PivotDistinct_NoDataModel
{
    class Program
    {
        static void Main(string[] args)
        {
            // Sample data - replace with your real DataTable
            DataTable dt = GetSampleDataTable();

            CreatePivotReportFromDistinct(dt,
                outputFolder: @"C:\Reports\Pivot\",
                baseFileName: "Monthly Application Onboarding Report");
        }

        public static void CreatePivotReportFromDistinct(DataTable dt, string outputFolder, string baseFileName)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                Console.WriteLine("DataTable is null or empty.");
                return;
            }

            // Config
            string dateCol = "AppliedOn";
            string appIdCol = "ApplicationID";
            string appNameCol = "ApplicationName";
            string categoryCol = "Category";
            string monthKeyCol = "ApplicationMonth"; // helper column
            string dataSheetName = "Data";
            string distinctSheetName = "DistinctData";
            string reportSheetName = "Report";
            string pivotName = "MonthlyPivot";
            string pivotStartCell = "A7";
            string titleCell = "A5";
            string titleText = "Monthly Application Onboarding Report";
            string footerText = "Internal";

            Directory.CreateDirectory(outputFolder);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputFile = Path.Combine(outputFolder, $"{baseFileName}_{timestamp}.xlsx");

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet dataSheet = null;
            Excel.Worksheet distinctSheet = null;
            Excel.Worksheet reportSheet = null;

            try
            {
                // 1) Add helper column ApplicationMonth to a copy of the DataTable
                DataTable dt2 = dt.Copy();
                if (!dt2.Columns.Contains(monthKeyCol))
                    dt2.Columns.Add(monthKeyCol, typeof(string));

                for (int r = 0; r < dt2.Rows.Count; r++)
                {
                    var raw = dt2.Rows[r][dateCol];
                    if (raw == DBNull.Value || raw == null)
                    {
                        dt2.Rows[r][monthKeyCol] = "";
                        continue;
                    }
                    DateTime d;
                    if (raw is DateTime) d = (DateTime)raw;
                    else if (!DateTime.TryParse(raw.ToString(), out d)) d = DateTime.MinValue;

                    dt2.Rows[r][monthKeyCol] = d == DateTime.MinValue ? "" : d.ToString("MMM-yyyy"); // e.g., Jan-2025
                }

                // 2) Build a distinct table with only the columns we need to determine uniqueness
                // Distinct by ApplicationMonth + ApplicationID + ApplicationName + Category
                DataView dv = new DataView(dt2);
                DataTable dtDistinct = dv.ToTable(true, monthKeyCol, appIdCol, appNameCol, categoryCol);

                // 3) Start Excel
                excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
                workbook = excelApp.Workbooks.Add();

                // 4) Write original dt2 to "Data" sheet (optional; helpful for raw records)
                dataSheet = (Excel.Worksheet)workbook.Worksheets[1];
                dataSheet.Name = dataSheetName;
                WriteDataTableToSheet(dt2, dataSheet, headerRow: 3, dataStartRow: 4);

                // 5) Write distinct table to "DistinctData" sheet (pivot source)
                distinctSheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                distinctSheet.Name = distinctSheetName;
                WriteDataTableToSheet(dtDistinct, distinctSheet, headerRow: 1, dataStartRow: 2);

                // 6) Create Report sheet with Title
                reportSheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                reportSheet.Name = reportSheetName;
                reportSheet.Range[titleCell].Value2 = titleText;
                reportSheet.Range[titleCell].Font.Bold = true;
                reportSheet.Range[titleCell].Font.Size = 14;
                reportSheet.Range[titleCell + ":" + "C5"].Merge();
                reportSheet.Range[titleCell + ":" + "C5"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // 7) Create PivotCache from the DistinctData range (regular pivot, xlDatabase)
                int distinctCols = dtDistinct.Columns.Count;
                int distinctRows = dtDistinct.Rows.Count;
                string lastColLetter = GetExcelColumnName(distinctCols);
                string distinctRangeAddress = $"A1:{lastColLetter}{1 + distinctRows}"; // header at row1
                Excel.Range distinctRange = distinctSheet.Range[distinctRangeAddress];

                Excel.PivotCaches pivotCaches = workbook.PivotCaches();
                Excel.PivotCache pc = pivotCaches.Create(
                    Excel.XlPivotTableSourceType.xlDatabase,
                    distinctRange,
                    Excel.XlPivotTableVersionList.xlPivotTableVersion15);

                // 8) Create PivotTable on Report sheet at A7
                Excel.Range pivotDest = reportSheet.Range[pivotStartCell];
                Excel.PivotTable pt = pc.CreatePivotTable(pivotDest, pivotName, Type.Missing, Excel.XlPivotTableVersionList.xlPivotTableVersion15);

                // 9) Configure pivot: rows ApplicationMonth -> ApplicationName -> Category
                Excel.PivotField pfMonth = (Excel.PivotField)pt.PivotFields(monthKeyCol);
                pfMonth.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfMonth.Position = 1;
                pfMonth.Caption = "Application Month"; // custom row header

                Excel.PivotField pfAppName = (Excel.PivotField)pt.PivotFields(appNameCol);
                pfAppName.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfAppName.Position = 2;

                Excel.PivotField pfCategory = (Excel.PivotField)pt.PivotFields(categoryCol);
                pfCategory.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfCategory.Position = 3;

                // 10) Values: Count of ApplicationID (on distinct dataset this yields distinct count)
                Excel.PivotField pfValues = (Excel.PivotField)pt.AddDataField(pt.PivotFields(appIdCol), "Distinct Applications", Excel.XlConsolidationFunction.xlCount);
                try { pfValues.NumberFormat = "#,##0"; } catch { }

                // 11) Layout & formatting
                pt.RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow);
                pt.DisplayFieldCaptions = true;
                pt.ShowDrillIndicators = true;
                pt.ColumnGrand = true;
                pt.RowGrand = true;
                try { pt.TableStyle2 = "PivotStyleLight16"; } catch { }

                // hide subtotals on category
                try
                {
                    pfCategory.Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
                }
                catch { }

                // indent category if supported
                try { pfCategory.Indent = 2; } catch { }

                // repeat labels ON if available
                try { pt.RepeatAllLabels(Excel.XlPivotFieldOrientation.xlRowField); } catch { }

                reportSheet.Columns.AutoFit();

                // 12) Footer "Internal" below pivot (F1 style: last used row + 2)
                int lastUsed = reportSheet.UsedRange.Rows.Count;
                int footerRow = lastUsed + 2;
                reportSheet.Cells[footerRow, 1] = footerText;
                ((Excel.Range)reportSheet.Cells[footerRow, 1]).Font.Italic = true;
                ((Excel.Range)reportSheet.Cells[footerRow, 1]).Font.Size = 10;

                // 13) Save workbook
                // remove file if exists
                if (File.Exists(outputFile)) File.Delete(outputFile);
                workbook.SaveAs(outputFile);
                Console.WriteLine("Saved: " + outputFile);

                // 14) Optionally make visible
                // excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
                throw;
            }
            finally
            {
                // Cleanup COM objects
                try { if (reportSheet != null) Marshal.FinalReleaseComObject(reportSheet); } catch { }
                try { if (distinctSheet != null) Marshal.FinalReleaseComObject(distinctSheet); } catch { }
                try { if (dataSheet != null) Marshal.FinalReleaseComObject(dataSheet); } catch { }

                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(true);
                        Marshal.FinalReleaseComObject(workbook);
                        workbook = null;
                    }
                }
                catch { }

                try
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.FinalReleaseComObject(excelApp);
                        excelApp = null;
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Helper: write DataTable to a worksheet fast (headers + data)
        private static void WriteDataTableToSheet(DataTable dt, Excel.Worksheet ws, int headerRow, int dataStartRow)
        {
            int cols = dt.Columns.Count;
            int rows = dt.Rows.Count;

            // headers
            for (int c = 0; c < cols; c++)
            {
                ws.Cells[headerRow, c + 1] = dt.Columns[c].ColumnName;
                ((Excel.Range)ws.Cells[headerRow, c + 1]).Font.Bold = true;
            }

            // build array
            object[,] arr = new object[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    arr[r, c] = dt.Rows[r][c] == DBNull.Value ? null : dt.Rows[r][c];

            Excel.Range startCell = ws.Cells[dataStartRow, 1];
            Excel.Range endCell = ws.Cells[dataStartRow + rows - 1, cols];
            Excel.Range writeRange = ws.Range[startCell, endCell];
            writeRange.Value2 = arr;

            // autofit
            ws.Columns.AutoFit();
        }

        // Helper: convert index -> Excel column (1 => A)
        private static string GetExcelColumnName(int columnNumber)
        {
            if (columnNumber < 1) throw new ArgumentOutOfRangeException(nameof(columnNumber));
            string columnName = String.Empty;
            while (columnNumber > 0)
            {
                columnNumber--;
                int remainder = columnNumber % 26;
                columnName = (char)('A' + remainder) + columnName;
                columnNumber = columnNumber / 26;
            }
            return columnName;
        }

        // Sample test data
        private static DataTable GetSampleDataTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("ApplicationID", typeof(string));
            dt.Columns.Add("AppliedOn", typeof(DateTime));
            dt.Columns.Add("ApplicationName", typeof(string));
            dt.Columns.Add("Category", typeof(string));

            dt.Rows.Add("App1", new DateTime(2025, 1, 2), "Alpha", "Finance");
            dt.Rows.Add("App2", new DateTime(2025, 1, 8), "Beta", "Retail");
            dt.Rows.Add("App1", new DateTime(2025, 1, 20), "Alpha", "Finance");
            dt.Rows.Add("App3", new DateTime(2025, 2, 5), "Gamma", "Retail");
            dt.Rows.Add("App4", new DateTime(2025, 2, 10), "Delta", "HR");
            dt.Rows.Add("App4", new DateTime(2025, 2, 15), "Delta", "HR");
            dt.Rows.Add("App5", new DateTime(2025, 3, 3), "Epsilon", "Finance");
            return dt;
        }
    }
}
