Attribute VB_Name = "BelarusianExample"
Option Explicit

' Прыклад апрацоўкі заказу на беларускай мове
Function РазлічыцьСуму(цана As Currency, колькасць As Long) As Currency
    If колькасць <= 0 Then Exit Function
    РазлічыцьСуму = цана * колькасць
End Function

Private Function ПраверыцьКліента(імя As String, тэлефон As String) As Boolean
    If Len(Trim$(імя)) = 0 Then Exit Function
    If Len(Trim$(тэлефон)) < 7 Then Exit Function
    ПраверыцьКліента = True
End Function

Private Function РазлічыцьЗніжку(сума As Currency, працэнт As Double) As Currency
    If сума <= 0 Or працэнт < 0 Or працэнт > 1 Then Exit Function
    РазлічыцьЗніжку = сума * працэнт
End Function

Sub ЗапоўніцьАркушЗамовы(аркуш As Worksheet)
    Dim радок As Long
    Dim сума As Currency
    аркуш.Range("A1:C1").Value = Array("Тавар", "Колькасць", "Сума")
    аркуш.Range("A1:C1").Font.Bold = True
    аркуш.Cells(2, 1).Value = "Кава і гарбата"
    аркуш.Cells(2, 2).Value = 4
    аркуш.Cells(2, 3).Value = 85.5 * аркуш.Cells(2, 2).Value
    аркуш.Cells(3, 1).Value = "Кніга на беларускай мове"
    аркуш.Cells(3, 2).Value = 2
    аркуш.Cells(3, 3).Value = 210 * аркуш.Cells(3, 2).Value
    радок = аркуш.Cells(аркуш.Rows.Count, 1).End(xlUp).Row
    сума = Application.Sum(аркуш.Range("C2:C" & радок))
    аркуш.Cells(радок, 4).Value = сума
    аркуш.Columns("A:D").AutoFit
End Sub

Sub ПаказацьВынік(імя As String, тэлефон As String)
    If ПраверыцьКліента(імя, тэлефон) Then
        MsgBox "Даныя кліента правераныя.", vbInformation, "Замова"
    Else
        MsgBox "Праверце імя і нумар тэлефона.", vbExclamation, "Памылка"
    End If
End Sub
