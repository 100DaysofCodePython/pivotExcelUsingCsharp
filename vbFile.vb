Sub TestPivot_DataModel_DistinctCount_Stable()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPT As Worksheet
    Dim lo As ListObject
    Dim conn As WorkbookConnection
    Dim pc As PivotCache
    Dim pt As PivotTable

    Set wb = ThisWorkbook

    ' 1) Create or refresh sample data
    On Error Resume Next
    Set wsData = wb.Sheets("SalesData")
    If wsData Is Nothing Then
        Set wsData = wb.Worksheets.Add
        wsData.Name = "SalesData"
    Else
        wsData.Cells.Clear
    End If
    On Error GoTo 0

    wsData.Range("A1:E1").Value = Array("Date", "CustomerID", "Category", "Month", "Amount")
    wsData.Range("A2:E7").Value = Array( _
        Array("2025-01-01", 101, "A", "Jan", 120), _
        Array("2025-01-02", 102, "A", "Jan", 140), _
        Array("2025-01-05", 101, "B", "Jan", 200), _
        Array("2025-02-03", 103, "A", "Feb", 190), _
        Array("2025-02-10", 102, "B", "Feb", 300), _
        Array("2025-02-15", 101, "A", "Feb", 250) _
    )

    ' 2) Convert to table if needed
    If wsData.ListObjects.Count = 0 Then
        wsData.ListObjects.Add xlSrcRange, wsData.Range("A1").CurrentRegion, , xlYes
    End If
    Set lo = wsData.ListObjects(1)
    lo.Name = "SalesTable"

    ' 3) Create a Data Model connection (THIS is the stable method)
    On Error Resume Next
    Set conn = wb.Connections("SalesTable_Connection")
    On Error GoTo 0

    If conn Is Nothing Then
        Set conn = wb.Connections.Add2( _
            Name:="SalesTable_Connection", _
            Description:="Push table to Data Model", _
            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & lo.Name & ";", _
            CommandText:=lo.Name, _
            lCmdtype:=xlCmdTable)
    End If

    ' 4) Now create PivotCache from this connection
    Set pc = wb.PivotCaches.Create( _
                SourceType:=xlExternal, _
                SourceData:="SalesTable_Connection")

    ' 5) Create pivot sheet
    On Error Resume Next
    Set wsPT = wb.Sheets("SalesPivot")
    If wsPT Is Nothing Then
        Set wsPT = wb.Worksheets.Add
        wsPT.Name = "SalesPivot"
    Else
        wsPT.Cells.Clear
    End If
    On Error GoTo 0

    ' 6) Create pivot
    Set pt = wsPT.PivotTables.Add(pc, wsPT.Range("A3"), "SalesPivotTable")

    ' 7) Add fields
    With pt
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Category").Orientation = xlColumnField

        Dim pf As PivotField
        Set pf = .PivotFields("CustomerID")
        pf.Orientation = xlDataField
        pf.Function = xlDistinctCount
        pf.Name = "Distinct Customers"
    End With

    MsgBox "✅ STABLE pivot with distinct count created successfully!"

End Sub
