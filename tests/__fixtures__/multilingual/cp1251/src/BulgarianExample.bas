Attribute VB_Name = "BulgarianExample"
Option Explicit

' Пример за обработка на поръчка на български език
Function ИзчислиСума(цена As Currency, количество As Long) As Currency
    If количество <= 0 Then Exit Function
    ИзчислиСума = цена * количество
End Function

Private Function ПровериКлиент(име As String, телефон As String) As Boolean
    If Len(Trim$(име)) = 0 Then Exit Function
    If Len(Trim$(телефон)) < 7 Then Exit Function
    ПровериКлиент = True
End Function

Private Function ИзчислиОтстъпка(сума As Currency, процент As Double) As Currency
    If сума <= 0 Or процент < 0 Or процент > 1 Then Exit Function
    ИзчислиОтстъпка = сума * процент
End Function

Sub ПопълниЛистПоръчка(лист As Worksheet)
    Dim ред As Long
    Dim сума As Currency
    лист.Range("A1:C1").Value = Array("Продукт", "Количество", "Сума")
    лист.Range("A1:C1").Font.Bold = True
    лист.Cells(2, 1).Value = "Кафе и чай"
    лист.Cells(2, 2).Value = 4
    лист.Cells(2, 3).Value = 85.5 * лист.Cells(2, 2).Value
    лист.Cells(3, 1).Value = "Книга на български език"
    лист.Cells(3, 2).Value = 2
    лист.Cells(3, 3).Value = 210 * лист.Cells(3, 2).Value
    ред = лист.Cells(лист.Rows.Count, 1).End(xlUp).Row
    сума = Application.Sum(лист.Range("C2:C" & ред))
    лист.Cells(ред, 4).Value = сума
    лист.Columns("A:D").AutoFit
End Sub

Sub ПокажиРезултат(име As String, телефон As String)
    If ПровериКлиент(име, телефон) Then
        MsgBox "Данните на клиента са проверени.", vbInformation, "Поръчка"
    Else
        MsgBox "Проверете името и телефонния номер.", vbExclamation, "Грешка"
    End If
End Sub
