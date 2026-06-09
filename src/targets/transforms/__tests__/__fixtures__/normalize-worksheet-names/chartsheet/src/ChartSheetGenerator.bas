Attribute VB_Name = "ChartSheetGenerator"
' Creates 2 chart sheets, then reorders them to test whether
' Excel swaps the file contents on disk (as it does with worksheets).
' Run ONCE after `vba build`, then `vba export` to capture the OOXML.
'
' @TODO Extend normalizeWorksheetNames (or add a separate
' normalizeChartsheetNames) to also rename chart sheet files to
' chrt{codeName}.xml and their _rels sidecars, preventing the same
' swap-on-reorder diff noise that was fixed for worksheets in #57.
Public Sub Main()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Add some data for the charts
    ws.Range("A1").Value = "Category"
    ws.Range("B1").Value = "Value"
    ws.Range("A2").Value = "A"
    ws.Range("B2").Value = 10
    ws.Range("A3").Value = "B"
    ws.Range("B3").Value = 20

    ' Create first chart sheet
    Dim chart1 As Chart
    Set chart1 = ThisWorkbook.Charts.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    chart1.ChartType = xlColumnClustered
    chart1.SetSourceData Source:=ws.Range("A1:B2")
    chart1.Name = "Chart A"

    ' Create second chart sheet
    Dim Chart2 As Chart
    Set Chart2 = ThisWorkbook.Charts.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    Chart2.ChartType = xlPie
    Chart2.SetSourceData Source:=ws.Range("A1:B3")
    Chart2.Name = "Chart B"

    ' Reorder: move "Chart B" BEFORE "Chart A" (swap tab order)
    Chart2.Move Before:=chart1

    ThisWorkbook.Save
End Sub
