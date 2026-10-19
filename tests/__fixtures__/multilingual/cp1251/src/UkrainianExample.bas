Attribute VB_Name = "UkrainianExample"
Option Explicit

' Приклад обробки замовлення українською мовою
Function ОбчислитиСуму(ціна As Currency, кількість As Long) As Currency
    If кількість <= 0 Then Exit Function
    ОбчислитиСуму = ціна * кількість
End Function

Private Function СтворитиРядокЗамовлення(назва As String, ціна As Currency, кількість As Long) As String
    If Len(Trim$(назва)) = 0 Or кількість <= 0 Then Exit Function
    СтворитиРядокЗамовлення = назва & " | " & ціна * кількість
End Function

Private Function ПеревіритиКлієнта(імя As String, телефон As String) As Boolean
    If Len(Trim$(імя)) = 0 Then Exit Function
    If Len(Trim$(телефон)) < 7 Then Exit Function
    ПеревіритиКлієнта = True
End Function

Sub ЗаповнитиАркушЗамовлення(аркуш As Worksheet)
    Dim рядок As Long
    Dim назва As String
    Dim ціна As Currency
    Dim кількість As Long
    аркуш.Range("A1:C1").Value = Array("Товар", "Кількість", "Сума")
    аркуш.Range("A1:C1").Font.Bold = True
    назва = "Кава та чай"
    ціна = 85.5
    кількість = 4
    рядок = 2
    аркуш.Cells(рядок, 1).Value = назва
    аркуш.Cells(рядок, 2).Value = кількість
    аркуш.Cells(рядок, 3).Value = ціна * кількість
    назва = "Книжка українською"
    ціна = 210
    кількість = 2
    рядок = рядок + 1
    аркуш.Cells(рядок, 1).Value = назва
    аркуш.Cells(рядок, 2).Value = кількість
    аркуш.Cells(рядок, 3).Value = ціна * кількість
    аркуш.Columns("A:C").AutoFit
End Sub

Sub ПоказатиРезультат(імя As String, телефон As String)
    If ПеревіритиКлієнта(імя, телефон) Then
        MsgBox "Дані клієнта перевірено.", vbInformation, "Замовлення"
    Else
        MsgBox "Перевірте ім'я та номер телефону.", vbExclamation, "Помилка"
    End If
End Sub
