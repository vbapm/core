Attribute VB_Name = "MacedonianExample"
Option Explicit

' Пример за обработка нарачка на македонски јазик
Function ПресметајСума(цена As Currency, количина As Long) As Currency
    If количина <= 0 Then Exit Function
    ПресметајСума = цена * количина
End Function

Private Function ПровериКлиент(име As String, телефон As String) As Boolean
    If Len(Trim$(име)) = 0 Then Exit Function
    If Len(Trim$(телефон)) < 7 Then Exit Function
    ПровериКлиент = True
End Function

Private Function ПресметајПопуст(сума As Currency, процент As Double) As Currency
    If сума <= 0 Or процент < 0 Or процент > 1 Then Exit Function
    ПресметајПопуст = сума * процент
End Function

Sub ПополниЛистНарачка(лист As Worksheet)
    Dim ред As Long
    Dim сума As Currency
    лист.Range("A1:C1").Value = Array("Производ", "Количина", "Износ")
    лист.Range("A1:C1").Font.Bold = True
    лист.Cells(2, 1).Value = "Кафе и чај"
    лист.Cells(2, 2).Value = 4
    лист.Cells(2, 3).Value = 85.5 * лист.Cells(2, 2).Value
    лист.Cells(3, 1).Value = "Книга на македонски јазик"
    лист.Cells(3, 2).Value = 2
    лист.Cells(3, 3).Value = 210 * лист.Cells(3, 2).Value
    ред = лист.Cells(лист.Rows.Count, 1).End(xlUp).Row
    сума = Application.Sum(лист.Range("C2:C" & ред))
    лист.Cells(ред, 4).Value = сума
    лист.Columns("A:D").AutoFit
End Sub

Sub ПрикажиРезултат(име As String, телефон As String)
    If ПровериКлиент(име, телефон) Then
        MsgBox "Податоците за клиентот се проверени.", vbInformation, "Нарачка"
    Else
        MsgBox "Проверете го името и телефонскиот број.", vbExclamation, "Грешка"
    End If
End Sub
