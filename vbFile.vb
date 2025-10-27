Sub CreatePivotWithDistinctCount()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPT As Worksheet
    Dim rng As Range
    Dim pc As PivotCache
    Dim pt As PivotTable

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("SalesData") ' Change to your data sheet name
    Set rng = wsData.Range("A1").CurrentRegion

    '--- Create connection to Data Model ---
    Set pc = wb.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=rng, _
                Version:=xlPivotTableVersion15)

    'Important: Add to the Data Model
    pc.EnableRefresh = True
    pc.Refresh
    pc.SaveData = True
    pc.OLAP = True    ' this is what pushes to the Data Model

    '--- Add new sheet for pivot ---
    On Error Resume Next
    Set wsPT = wb.Sheets("SalesPivot")
    If wsPT Is Nothing Then
        Set wsPT = wb.Worksheets.Add
        wsPT.Name = "SalesPivot"
    End If
    On Error GoTo 0
    wsPT.Cells.Clear

    '--- Create pivot table ---
    Set pt = wsPT.PivotTables.Add( _
                PivotCache:=pc, _
                TableDestination:=wsPT.Range("A3"), _
                TableName:="SalesPivotTable")

    '--- Example fields (Change to your columns) ---
    With pt
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Category").Orientation = xlColumnField

        ' DISTINCT COUNT only works when Data Model connection is OLAP
        Dim pf As PivotField
        Set pf = .PivotFields("CustomerID")    ' example field

        pf.Orientation = xlDataField
        pf.Function = xlDistinctCount  ' this is the KEY line
        pf.Name = "Distinct Customers"
    End With

    MsgBox "Pivot with distinct count created successfully!", vbInformation

End Sub
