Attribute VB_Name = "SerbianExample"
Option Explicit

' Пример обраде поруџбине на српском језику
Function ИзрачунајСуму(цена As Currency, количина As Long) As Currency
    If количина <= 0 Then Exit Function
    ИзрачунајСуму = цена * количина
End Function

Private Function ПровериКупца(име As String, телефон As String) As Boolean
    If Len(Trim$(име)) = 0 Then Exit Function
    If Len(Trim$(телефон)) < 7 Then Exit Function
    ПровериКупца = True
End Function

Private Function ИзрачунајПопуст(сума As Currency, проценат As Double) As Currency
    If сума <= 0 Or проценат < 0 Or проценат > 1 Then Exit Function
    ИзрачунајПопуст = сума * проценат
End Function

Sub ПопуниЛистПоруџбина(лист As Worksheet)
    Dim ред As Long
    Dim сума As Currency
    лист.Range("A1:C1").Value = Array("Производ", "Количина", "Износ")
    лист.Range("A1:C1").Font.Bold = True
    лист.Cells(2, 1).Value = "Кафа и чај"
    лист.Cells(2, 2).Value = 4
    лист.Cells(2, 3).Value = 85.5 * лист.Cells(2, 2).Value
    лист.Cells(3, 1).Value = "Књига на српском језику"
    лист.Cells(3, 2).Value = 2
    лист.Cells(3, 3).Value = 210 * лист.Cells(3, 2).Value
    ред = лист.Cells(лист.Rows.Count, 1).End(xlUp).Row
    сума = Application.Sum(лист.Range("C2:C" & ред))
    лист.Cells(ред, 4).Value = сума
    лист.Columns("A:D").AutoFit
End Sub

Sub ПрикажиРезултат(име As String, телефон As String)
    If ПровериКупца(име, телефон) Then
        MsgBox "Подаци о купцу су проверени.", vbInformation, "Поруџбина"
    Else
        MsgBox "Проверите име и број телефона.", vbExclamation, "Грешка"
    End If
End Sub
