Sub CreatePivotWithDistinctCount_DataModel()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPT As Worksheet
    Dim lo As ListObject
    Dim pc As PivotCache
    Dim pt As PivotTable

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("SalesData") ' your sheet

    ' your range MUST be a ListObject (Excel Table) for Data Model
    If wsData.ListObjects.Count = 0 Then
        wsData.ListObjects.Add xlSrcRange, wsData.Range("A1").CurrentRegion, , xlYes
    End If

    Set lo = wsData.ListObjects(1)

    ' --- Create pivot cache directly with Data Model flag ---
    Set pc = wb.PivotCaches.Create( _
                SourceType:=xlExternal, _
                SourceData:=Array("WORKSHEET;" & wb.FullName), _
                Version:=xlPivotTableVersion15)

    pc.MaintainConnection = True
    pc.EnableRefresh = True

    ' add the ListObject to the connection
    pc.CommandType = xlCmdExcel
    pc.CommandText = lo.Name

    ' --- Create sheet for Pivot ---
    On Error Resume Next
    Set wsPT = wb.Worksheets("SalesPivot")
    If wsPT Is Nothing Then
        Set wsPT = wb.Worksheets.Add
        wsPT.Name = "SalesPivot"
    End If
    On Error GoTo 0
    wsPT.Cells.Clear

    ' --- Create PivotTable ---
    Set pt = wsPT.PivotTables.Add( _
             PivotCache:=pc, _
             TableDestination:=wsPT.Range("A3"), _
             TableName:="SalesPivotTable")

    ' --- Add fields ---
    With pt
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Category").Orientation = xlColumnField

        Dim pf As PivotField
        Set pf = .PivotFields("CustomerID")

        pf.Orientation = xlDataField
        pf.Function = xlDistinctCount     ' DISTINCT COUNT now appears ✅
        pf.Name = "Distinct Customers"
    End With

    MsgBox "Pivot with Data Model distinct count created!", vbInformation
End Sub
