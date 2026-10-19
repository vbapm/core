Attribute VB_Name = "MongolianExample"
Option Explicit

' CP1251 cannot represent every Mongolian Cyrillic letter.
' This sample uses the representable part of the Mongolian alphabet.
Function CalculateTotal(price As Currency, quantity As Long) As Currency
    If quantity <= 0 Then Exit Function
    CalculateTotal = price * quantity
End Function

Private Function CheckCustomer(name As String, phone As String) As Boolean
    If Len(Trim$(name)) = 0 Then Exit Function
    If Len(Trim$(phone)) < 7 Then Exit Function
    CheckCustomer = True
End Function

Private Function CalculateDiscount(total As Currency, rate As Double) As Currency
    If total <= 0 Or rate < 0 Or rate > 1 Then Exit Function
    CalculateDiscount = total * rate
End Function

Sub FillOrderSheet(sheet As Worksheet)
    Dim rowNumber As Long
    Dim total As Currency
    sheet.Range("A1:C1").Value = Array("Бараа", "Тоо", "Нийт")
    sheet.Range("A1:C1").Font.Bold = True
    sheet.Cells(2, 1).Value = "Кофе ба цай"
    sheet.Cells(2, 2).Value = 4
    sheet.Cells(2, 3).Value = 85.5 * sheet.Cells(2, 2).Value
    sheet.Cells(3, 1).Value = "Монгол ном"
    sheet.Cells(3, 2).Value = 2
    sheet.Cells(3, 3).Value = 210 * sheet.Cells(3, 2).Value
    rowNumber = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
    total = Application.Sum(sheet.Range("C2:C" & rowNumber))
    sheet.Cells(rowNumber, 4).Value = total
    sheet.Columns("A:D").AutoFit
End Sub

Sub ShowResult(name As String, phone As String)
    If CheckCustomer(name, phone) Then
        MsgBox "Харилцагчийн мэдээлэл шалгагдлаа.", vbInformation, "Захиалга"
    Else
        MsgBox "Нэр болон утасны дугаарыг шалгана уу.", vbExclamation, "Алдаа"
    End If
End Sub
