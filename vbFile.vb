Sub TestPivot_DataModel_DistinctCount()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPT As Worksheet
    Dim lo As ListObject
    Dim pc As PivotCache
    Dim pt As PivotTable

    Set wb = ThisWorkbook

    '==== 1. CREATE SAMPLE DATA SHEET ====
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

    '==== 2. MAKE SURE IT IS A TABLE ====
    If wsData.ListObjects.Count = 0 Then
        wsData.ListObjects.Add xlSrcRange, wsData.Range("A1").CurrentRegion, , xlYes
    End If

    Set lo = wsData.ListObjects(1)
    lo.Name = "SalesTable"

    '==== 3. CREATE PIVOT SHEET ====
    On Error Resume Next
    Set wsPT = wb.Sheets("SalesPivot")
    If wsPT Is Nothing Then
        Set wsPT = wb.Worksheets.Add
        wsPT.Name = "SalesPivot"
    Else
        wsPT.Cells.Clear
    End If
    On Error GoTo 0

    '==== 4. CREATE DATA MODEL PIVOT CACHE (KEY PART) ====
    Set pc = wb.PivotCaches.Create( _
                SourceType:=xlExternal, _
                SourceData:=Array("WORKSHEET;" & wb.FullName), _
                Version:=xlPivotTableVersion15)

    pc.CommandType = xlCmdExcel
    pc.CommandText = lo.Name
    pc.MaintainConnection = True

    '==== 5. CREATE PIVOT TABLE FROM DATA MODEL ====
    Set pt = wsPT.PivotTables.Add( _
                PivotCache:=pc, _
                TableDestination:=wsPT.Range("A3"), _
                TableName:="SalesPivotTable")

    '==== 6. ADD FIELDS ====
    With pt
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Category").Orientation = xlColumnField

        Dim pf As PivotField
        Set pf = .PivotFields("CustomerID")
        pf.Orientation = xlDataField
        pf.Function = xlDistinctCount
        pf.Name = "Distinct Customers"
    End With

    MsgBox "✅ Pivot with Data Model distinct count created successfully!"

End Sub
