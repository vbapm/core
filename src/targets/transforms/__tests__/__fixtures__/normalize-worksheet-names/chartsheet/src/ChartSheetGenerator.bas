Attribute VB_Name = "ChartSheetGenerator"
' Adds a chart sheet to the workbook.
' Run this ONCE after `vba build`, then `vba export` to capture the OOXML.
Public Sub Main()
    Dim chartSheet As Chart
    Set chartSheet = ThisWorkbook.Charts.Add

    ' Add a simple chart to the sheet so there's actual content
    chartSheet.ChartType = xlColumnClustered
    chartSheet.SetSourceData Source:=Sheets("Sheet1").Range("A1:B2")

    ThisWorkbook.Save
End Sub
