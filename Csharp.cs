// Requires reference to Microsoft.Office.Interop.Excel
// Target .NET Framework (e.g., 4.7.2 or 4.8). Run on Windows with Excel installed (x64 recommended).

using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataModelPivotFinal
{
    class Program
    {
        static void Main(string[] args)
        {
            // Build or get your DataTable here. Replace GetSampleDataTable() with real source.
            DataTable dt = GetSampleDataTable();
            CreatePivotReport_DataModel(dt,
                outputFolder: @"C:\Reports\Pivot\",
                outputBaseFileName: "Monthly Application Onboarding Report",
                tableName: "ApplicationsTbl"   // <-- user-chosen table name (one-word)
            );
        }

        public static void CreatePivotReport_DataModel(DataTable dt, string outputFolder, string outputBaseFileName, string tableName)
        {
            if (dt == null || dt.Rows.Count == 0) throw new ArgumentException("DataTable is null or empty.");

            // Configs (adjust if needed)
            string dateCol = "AppliedOn";
            string appIdCol = "ApplicationID";
            string appNameCol = "ApplicationName";
            string categoryCol = "Category";
            string monthKeyCol = "ApplicationMonth"; // helper column
            string reportSheetName = "Report";
            string dataSheetName = "Data";
            string pivotName = "MonthlyPivot";
            string pivotStartCell = "A7";
            string titleCell = "A5";
            string titleText = "Monthly Application Onboarding Report";
            string footerText = "Internal";

            Directory.CreateDirectory(outputFolder);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputFile = Path.Combine(outputFolder, $"{outputBaseFileName}_{timestamp}.xlsx");

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet dataSheet = null;
            Excel.Worksheet reportSheet = null;

            try
            {
                // 1) Prepare DataTable copy and add month helper column (string "MMM-yyyy")
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
                    dt2.Rows[r][monthKeyCol] = d == DateTime.MinValue ? "" : d.ToString("MMM-yyyy");
                }

                // 2) Start Excel
                excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                workbook = excelApp.Workbooks.Add();
                dataSheet = (Excel.Worksheet)workbook.Worksheets[1];
                dataSheet.Name = dataSheetName;

                // 3) Write headers at row 3 and data from row 4 in one shot (fast)
                int headerRow = 3;
                int dataStartRow = headerRow + 1;
                int cols = dt2.Columns.Count;
                int rows = dt2.Rows.Count;

                for (int c = 0; c < cols; c++)
                {
                    dataSheet.Cells[headerRow, c + 1] = dt2.Columns[c].ColumnName;
                    ((Excel.Range)dataSheet.Cells[headerRow, c + 1]).Font.Bold = true;
                }

                object[,] arr = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        arr[r, c] = dt2.Rows[r][c] == DBNull.Value ? null : dt2.Rows[r][c];

                Excel.Range startCell = (Excel.Range)dataSheet.Cells[dataStartRow, 1];
                Excel.Range endCell = (Excel.Range)dataSheet.Cells[headerRow + rows, cols];
                Excel.Range writeRange = dataSheet.Range[startCell, endCell];
                writeRange.Value2 = arr;
                dataSheet.Columns.AutoFit();

                // 4) Create ListObject (Excel Table) for the range (header + data)
                string lastCol = GetExcelColumnName(cols);
                string tableAddress = $"A{headerRow}:{lastCol}{headerRow + rows}";
                Excel.Range tableRange = dataSheet.Range[tableAddress];

                Excel.ListObject listObj = dataSheet.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    tableRange,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing);

                // ensure table has the requested name, or append timestamp if exists
                try { listObj.Name = tableName; } catch { listObj.Name = tableName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss"); }

                // 5) Save workbook first — important so workbook.FullName exists for connection
                workbook.SaveAs(outputFile, Excel.XlFileFormat.xlOpenXMLWorkbook);

                // 6) Create a Workbook Connection referencing the table and promote to Data Model
                Excel.WorkbookConnections connections = workbook.Connections;
                string connName = "Conn_" + tableName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                Excel.WorkbookConnection conn = null;

                try
                {
                    // Add2 is preferred
                    conn = connections.Add2(
                        connName,
                        "Data Model connection",
                        $"WORKSHEET;{workbook.FullName}",
                        listObj.Name,
                        (int)Excel.XlCmdType.xlCmdExcel);
                }
                catch
                {
                    // fallback to Add
                    conn = connections.Add(connName, "Data Model connection (fallback)", $"WORKSHEET;{workbook.FullName}", listObj.Name);
                }

                // Try to promote to data model by setting ModelTableName
                try
                {
                    conn.ModelTableName = listObj.Name;
                }
                catch
                {
                    // some builds do not allow direct setting; it's okay — Excel may still add the model table
                }

                // Small pause can help Excel process the connection
                System.Threading.Thread.Sleep(400);

                // 7) Create PivotCache as external using the connection. Use conn object or name
                Excel.PivotCaches pivotCaches = workbook.PivotCaches();
                Excel.PivotCache pc = null;
                try
                {
                    pc = pivotCaches.Create(Excel.XlPivotTableSourceType.xlExternal, conn, Excel.XlPivotTableVersionList.xlPivotTableVersion15);
                }
                catch
                {
                    // fallback: pass connection name
                    try
                    {
                        pc = pivotCaches.Create(Excel.XlPivotTableSourceType.xlExternal, conn.Name, Excel.XlPivotTableVersionList.xlPivotTableVersion15);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Could not create PivotCache from connection. " + ex.Message);
                    }
                }

                // 8) Create Report sheet and place pivot at A7
                reportSheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                reportSheet.Name = reportSheetName;

                reportSheet.Range[titleCell].Value2 = titleText;
                reportSheet.Range[titleCell].Font.Bold = true;
                reportSheet.Range[titleCell].Font.Size = 14;
                reportSheet.Range[titleCell + ":" + "C5"].Merge();
                reportSheet.Range[titleCell + ":" + "C5"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range pivotDest = reportSheet.Range[pivotStartCell];

                Excel.PivotTable pt = pc.CreatePivotTable(pivotDest, pivotName, Type.Missing, Excel.XlPivotTableVersionList.xlPivotTableVersion15);

                // 9) Configure Pivot fields:
                // Row1: monthKeyCol
                Excel.PivotField pfMonth = (Excel.PivotField)pt.PivotFields(monthKeyCol);
                pfMonth.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfMonth.Position = 1;
                pfMonth.Caption = "Application Month"; // user requested caption

                // Row2: ApplicationName
                Excel.PivotField pfAppName = (Excel.PivotField)pt.PivotFields(appNameCol);
                pfAppName.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfAppName.Position = 2;

                // Row3: Category
                Excel.PivotField pfCategory = (Excel.PivotField)pt.PivotFields(categoryCol);
                pfCategory.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pfCategory.Position = 3;

                // 10) Add Distinct Count on ApplicationID (requires Data Model)
                Excel.PivotField pfData = (Excel.PivotField)pt.AddDataField(pt.PivotFields(appIdCol), "Distinct Applications", Excel.XlConsolidationFunction.xlDistinctCount);
                try { pfData.NumberFormat = "#,##0"; } catch { }

                // 11) Layout & cosmetics
                pt.RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow);
                pt.DisplayFieldCaptions = true;
                pt.ShowDrillIndicators = true;
                pt.ColumnGrand = true;
                pt.RowGrand = true;
                try { pt.TableStyle2 = "PivotStyleLight16"; } catch { }

                // Hide subtotals for category level
                try
                {
                    pfCategory.Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
                }
                catch { }

                // Indent category (if supported)
                try { pfCategory.Indent = 2; } catch { }

                // Repeat labels (if supported)
                try { pt.RepeatAllLabels(Excel.XlPivotFieldOrientation.xlRowField); } catch { }

                // Autofit report sheet columns
                reportSheet.Columns.AutoFit();

                // Footer placement (F1): last used row + 2
                int lastUsed = reportSheet.UsedRange.Rows.Count;
                int footerRow = lastUsed + 2;
                reportSheet.Cells[footerRow, 1] = footerText;
                ((Excel.Range)reportSheet.Cells[footerRow, 1]).Font.Italic = true;
                ((Excel.Range)reportSheet.Cells[footerRow, 1]).Font.Size = 10;

                // Refresh pivot and set refresh on open (best-effort)
                try
                {
                    pt.RefreshTable();
                    workbook.RefreshAll();
                }
                catch { }

                // Save final workbook (ensures pivot parts persisted)
                workbook.Save();

                Console.WriteLine("Saved file: " + outputFile);
            }
            catch (COMException comEx)
            {
                Console.WriteLine("COM exception: " + comEx.Message);
                Console.WriteLine(comEx.StackTrace);
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
                throw;
            }
            finally
            {
                // Cleanup COM objects carefully
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

        // Helper: convert 1-based column index to Excel column letters
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

        // Sample DataTable for quick testing
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
