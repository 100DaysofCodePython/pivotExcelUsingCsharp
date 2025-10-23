// Requires reference: Microsoft.Office.Interop.Excel
// Target .NET Framework (e.g., 4.7.2 or 4.8). Run on Windows with Excel installed (x64).

using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PivotDistinctFixed
{
    class Program
    {
        static void Main(string[] args)
        {
            // For quick testing use sample data. Replace with your real DataTable as needed.
            DataTable dt = GetSampleDataTable();

            CreatePivotReport(dt);
            Console.WriteLine("Done. Check output folder.");
        }

        public static void CreatePivotReport(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                Console.WriteLine("DT is null or empty.");
                return;
            }

            // === Configurable items ===
            string folder = @"C:\Reports\Pivot\";
            Directory.CreateDirectory(folder);
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputFile = Path.Combine(folder, $"Monthly Application Onboarding Report_{timeStamp}.xlsx");

            string dataSheetName = "Data";
            string reportSheetName = "Report";
            string tableName = "tblData";
            string pivotName = "MonthlyPivot";
            string pivotStartCell = "A7"; // on Report sheet
            string titleCell = "A5";
            string titleText = "Monthly Application Onboarding Report";
            string footerText = "Internal";

            // Column names in your DataTable (must match exactly)
            string dateCol = "AppliedOn";          // must be DateTime-compatible
            string appIdCol = "ApplicationID";     // distinct count on this
            string appNameCol = "ApplicationName";
            string categoryCol = "Category";
            string monthKeyCol = "ApplicationMonth"; // helper column we will add

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet dataSheet = null;
            Excel.Worksheet reportSheet = null;

            try
            {
                // 1) Prepare dt copy and add ApplicationMonth helper column (MMM-yyyy)
                DataTable dt2 = dt.Copy();
                if (!dt2.Columns.Contains(monthKeyCol))
                    dt2.Columns.Add(monthKeyCol, typeof(string));

                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    var v = dt2.Rows[i][dateCol];
                    if (v == DBNull.Value || v == null)
                    {
                        dt2.Rows[i][monthKeyCol] = "";
                        continue;
                    }
                    DateTime d;
                    if (v is DateTime) d = (DateTime)v;
                    else if (!DateTime.TryParse(v.ToString(), out d)) d = DateTime.MinValue;

                    dt2.Rows[i][monthKeyCol] = d == DateTime.MinValue ? "" : d.ToString("MMM-yyyy"); // "Jan-2025"
                }

                // 2) Start Excel and workbook
                excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
                workbook = excelApp.Workbooks.Add();
                dataSheet = (Excel.Worksheet)workbook.Worksheets[1];
                dataSheet.Name = dataSheetName;

                // 3) Write header row (row 3) and data starting row 4
                int headerRow = 3;
                int dataStartRow = 4;
                int cols = dt2.Columns.Count;
                int rows = dt2.Rows.Count;

                for (int c = 0; c < cols; c++)
                {
                    dataSheet.Cells[headerRow, c + 1] = dt2.Columns[c].ColumnName;
                    ((Excel.Range)dataSheet.Cells[headerRow, c + 1]).Font.Bold = true;
                }

                // 4) Build object[,] and write in one shot
                object[,] arr = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        arr[r, c] = dt2.Rows[r][c] == DBNull.Value ? null : dt2.Rows[r][c];

                Excel.Range startCell = (Excel.Range)dataSheet.Cells[dataStartRow, 1];
                Excel.Range endCell = (Excel.Range)dataSheet.Cells[headerRow + rows, cols];
                Excel.Range writeRange = dataSheet.Range[startCell, endCell];
                writeRange.Value2 = arr;
                dataSheet.Columns.AutoFit();

                // 5) Create a ListObject (Excel Table) covering header+data
                string lastColLetter = GetExcelColumnName(cols);
                string tableAddress = $"A{headerRow}:{lastColLetter}{headerRow + rows}";
                Excel.Range tableRange = dataSheet.Range[tableAddress];

                Excel.ListObject lo = dataSheet.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    tableRange,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing);

                // ensure unique table name
                try { lo.Name = tableName; } catch { lo.Name = tableName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss"); }

                // 6) Save workbook first (important — workbook.FullName will be populated)
                // Save as XLSX
                workbook.SaveAs(outputFile, Excel.XlFileFormat.xlOpenXMLWorkbook);
                // Now workbook.FullName is valid and file exists on disk.

                // 7) Create a workbook connection pointing to that table and promote to Data Model
                Excel.WorkbookConnections connections = workbook.Connections;
                string connName = "DataModelConnection_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                Excel.WorkbookConnection conn = null;
                try
                {
                    conn = connections.Add2(
                        connName,
                        "Connection for Data Model (table)",
                        $"WORKSHEET;{workbook.FullName}",
                        lo.Name,
                        (int)Excel.XlCmdType.xlCmdExcel);
                }
                catch
                {
                    // fallback to Add if Add2 not available
                    conn = connections.Add(connName, "Connection (fallback)", $"WORKSHEET;{workbook.FullName}", lo.Name);
                }

                // Attempt to set ModelTableName - if permitted this helps promote to the Data Model
                try { conn.ModelTableName = lo.Name; } catch { /* ignore if not supported */ }

                // 8) Create a PivotCache as external using the connection name/object
                Excel.PivotCaches pivotCaches = workbook.PivotCaches();
                Excel.PivotCache pc = null;
                try
                {
                    // Preferred: use the connection object
                    pc = pivotCaches.Create(Excel.XlPivotTableSourceType.xlExternal, conn, Excel.XlPivotTableVersionList.xlPivotTableVersion15);
                }
                catch
                {
                    // Fallback: pass connection name string if object fails
                    pc = pivotCaches.Create(Excel.XlPivotTableSourceType.xlExternal, conn.Name, Excel.XlPivotTableVersionList.xlPivotTableVersion15);
                }

                // 9) Create Report sheet and title
                reportSheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                reportSheet.Name = reportSheetName;

                // Title at A5
                reportSheet.Range[titleCell].Value2 = titleText;
                reportSheet.Range[titleCell].Font.Bold = true;
                reportSheet.Range[titleCell].Font.Size = 14;
                reportSheet.Range[titleCell + ":" + "C5"].Merge();
                reportSheet.Range[titleCell + ":" + "C5"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // 10) Create pivot table on Report sheet at A7
                Excel.Range pivotDest = reportSheet.Range[pivotStartCell];
                Excel.PivotTable pt = pc.CreatePivotTable(pivotDest, pivotName, Type.Missing, Excel.XlPivotTableVersionList.xlPivotTableVersion15);

                // 11) Configure pivot fields and hierarchy
                // Row1: ApplicationMonth (we created month string)
                Excel.PivotField pfMonth = (Excel.PivotField)pt.PivotFields(monthKeyCol);
                pfMonth.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfMonth.Position = 1;
                pfMonth.Caption = "Application Month"; // as requested

                // Row2: ApplicationName
                Excel.PivotField pfAppName = (Excel.PivotField)pt.PivotFields(appNameCol);
                pfAppName.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfAppName.Position = 2;

                // Row3: Category
                Excel.PivotField pfCategory = (Excel.PivotField)pt.PivotFields(categoryCol);
                pfCategory.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfCategory.Position = 3;

                // Values: Distinct Count on ApplicationID
                Excel.PivotField dataField = (Excel.PivotField)pt.AddDataField(pt.PivotFields(appIdCol), "Distinct Applications", Excel.XlConsolidationFunction.xlDistinctCount);
                try { dataField.NumberFormat = "#,##0"; } catch { }

                // 12) Layout and formatting
                pt.RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow);
                pt.DisplayFieldCaptions = true;
                pt.ShowDrillIndicators = true;
                pt.ColumnGrand = true;
                pt.RowGrand = true;
                try { pt.TableStyle2 = "PivotStyleLight16"; } catch { }

                // Attempt some Field formatting
                try
                {
                    // Hide subtotal for category
                    pfCategory.Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
                }
                catch { }

                // Indent (if supported)
                try { pfCategory.Indent = 2; } catch { }

                // Repeat labels ON (if available)
                try { pt.RepeatAllLabels(Excel.XlPivotFieldOrientation.xlRowField); } catch { }

                // Autofit report sheet columns
                reportSheet.Columns.AutoFit();

                // 13) Footer placement (F1 = last used row + 2)
                int usedRows = reportSheet.UsedRange.Rows.Count;
                int footerRow = usedRows + 2;
                reportSheet.Cells[footerRow, 1] = footerText;
                ((Excel.Range)reportSheet.Cells[footerRow, 1]).Font.Italic = true;
                ((Excel.Range)reportSheet.Cells[footerRow, 1]).Font.Size = 10;

                // 14) Refresh pivot and set refresh on open
                try
                {
                    pt.RefreshTable();
                    // Set workbook to refresh all connections on open
                    workbook.RefreshAll();
                }
                catch { }

                // 15) Save workbook (already saved earlier but final Save to ensure pivot parts are persisted)
                workbook.Save();

                Console.WriteLine("File saved to: " + outputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Clean up COM objects
                try { if (reportSheet != null) Marshal.FinalReleaseComObject(reportSheet); } catch { }
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

        // Helper - Excel column label
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

        // Sample data builder
        private static DataTable GetSampleDataTable()
        {
            DataTable dt = new DataTable();
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
