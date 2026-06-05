Attribute VB_Name = "Generator"
' Adds a simple shape to Sheet1 so the workbook has a drawing relationship.
' Run this ONCE after `vba build`, then `vba export` to capture the OOXML.
Public Sub Main()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Remove any existing shapes to keep the result deterministic
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' Add a simple rectangle -- produces xl/drawings/drawing1.xml +
    ' xl/worksheets/_rels/sheet1.xml.rels in the exported OOXML
    Dim newShp As Shape
    Set newShp = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 100, 30)
    newShp.Name = "MyShape"

    ThisWorkbook.Save
End Sub
