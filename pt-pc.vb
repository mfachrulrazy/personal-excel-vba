Option Explicit

Sub BuatPT()

Dim pc As PivotCache
Dim pt As PivotTable
Dim pf As PivotField
Dim ws As Worksheet
Dim wsg As Integer
Dim i As Integer

wsg = Worksheets.Count

For i = wsg To 2 Step -1
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="'" & Worksheets(i).Name & "'!" & Worksheets(i).Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1))
        
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "day" & i - 1
    Range("A5").Select
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ActiveCell, _
        TableName:="Day" & i - 1)
        
    
    pt.AddFields _
        RowFields:="Start Time", _
        ColumnFields:="Col1", _
        PageFields:=Array("Col2", "Col3", "Col4", "Col5")
    
    pt.AddDataField pt.PivotFields("Col6"), , xlAverage
Next i

End Sub

Sub buatChart()

Dim sh As Shape
Dim wsc As Worksheet
Dim ch As Chart

'Set wsc = Worksheets("day1")
For Each wsc In Worksheets
        If wsc.Name Like "day*" Then
            Set sh = wsc.Shapes.AddChart2( _
                XlChartType:=XlChartType.xlLineMarkers, _
                Left:=300, Top:=70, Width:=500, Height:=300)
            Set ch = sh.Chart
            ch.ClearToMatchStyle
            ch.ChartStyle = 233
        End If
Next

End Sub
