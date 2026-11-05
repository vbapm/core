Attribute VB_Name = "Utf8"
'===============================================================================
' Adapted from VBA StringTools
' https://github.com/guwidoe/VBA-StringTools
' MIT License
'
' Copyright (c) 2024 Guido Witt-Dörring
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'===============================================================================

Option Explicit
Option Base 0

Public Function ReadUtf8File(ByVal filePath As String) As String
    Dim encoded() As Byte
    Dim fileNumber As Integer
    Dim fileIsOpen As Boolean
    Dim byteCount As Long
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    On Error GoTo ErrorHandling

    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber
    fileIsOpen = True
    byteCount = LOF(fileNumber)
    If byteCount = 0 Then
        Close #fileNumber
        fileIsOpen = False
        ReadUtf8File = ""
        Exit Function
    End If

    ReDim encoded(0 To byteCount - 1)
    Get #fileNumber, 1, encoded
    Close #fileNumber
    fileIsOpen = False

    ReadUtf8File = DecodeUTF8(encoded)
    Exit Function

ErrorHandling:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    On Error Resume Next
    If fileIsOpen Then Close #fileNumber
    On Error GoTo 0
    Err.Raise savedErrorNumber, "Utf8.ReadUtf8File", savedErrorDescription
End Function

Public Function WriteUtf8File(ByVal filePath As String, ByVal text As String) As Boolean
    Dim encoded() As Byte
    Dim fileNumber As Integer
    Dim fileIsOpen As Boolean
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    On Error GoTo ErrorHandling

    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    fileIsOpen = True
    Close #fileNumber
    fileIsOpen = False

    Open filePath For Binary Access Write As #fileNumber
    fileIsOpen = True
    If Len(text) > 0 Then
        EncodeUTF8 text, encoded
        Put #fileNumber, 1, encoded
    End If
    Close #fileNumber
    fileIsOpen = False

    WriteUtf8File = True
    Exit Function

ErrorHandling:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    On Error Resume Next
    If fileIsOpen Then Close #fileNumber
    On Error GoTo 0
    Err.Raise savedErrorNumber, "Utf8.WriteUtf8File", savedErrorDescription
End Function

Private Function DecodeUTF8(ByRef utf8() As Byte) As String
    Dim byteCount As Long
    Dim inputPosition As Long
    Dim outputPosition As Long
    Dim firstByte As Long
    Dim secondByte As Long
    Dim thirdByte As Long
    Dim fourthByte As Long
    Dim codepoint As Long
    Dim outputText As String

    byteCount = UBound(utf8) + 1
    outputText = String$(byteCount, vbNullChar)
    If byteCount >= 3 Then
        If utf8(0) = &HEF And utf8(1) = &HBB And utf8(2) = &HBF Then inputPosition = 3
    End If
    outputPosition = 1

    Do While inputPosition < byteCount
        firstByte = CLng(utf8(inputPosition))

        If firstByte < &H80& Then
            codepoint = firstByte
            inputPosition = inputPosition + 1
        ElseIf firstByte >= &HC2& And firstByte <= &HDF& Then
            If inputPosition + 1 >= byteCount Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            secondByte = CLng(utf8(inputPosition + 1))
            If (secondByte And &HC0&) <> &H80& Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            codepoint = ((firstByte And &H1F&) * &H40&) Or (secondByte And &H3F&)
            inputPosition = inputPosition + 2
        ElseIf firstByte >= &HE0& And firstByte <= &HEF& Then
            If inputPosition + 2 >= byteCount Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            secondByte = CLng(utf8(inputPosition + 1))
            thirdByte = CLng(utf8(inputPosition + 2))
            If (secondByte And &HC0&) <> &H80& Or (thirdByte And &HC0&) <> &H80& Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            If firstByte = &HE0& And secondByte < &HA0& Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            If firstByte = &HED& And secondByte >= &HA0& Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            codepoint = ((firstByte And &HF&) * &H1000&) Or ((secondByte And &H3F&) * &H40&) Or (thirdByte And &H3F&)
            inputPosition = inputPosition + 3
        ElseIf firstByte >= &HF0& And firstByte <= &HF4& Then
            If inputPosition + 3 >= byteCount Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            secondByte = CLng(utf8(inputPosition + 1))
            thirdByte = CLng(utf8(inputPosition + 2))
            fourthByte = CLng(utf8(inputPosition + 3))
            If (secondByte And &HC0&) <> &H80& Or (thirdByte And &HC0&) <> &H80& Or (fourthByte And &HC0&) <> &H80& Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            If firstByte = &HF0& And secondByte < &H90& Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            If firstByte = &HF4& And secondByte > &H8F& Then Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
            codepoint = ((firstByte And &H7&) * &H40000) Or ((secondByte And &H3F&) * &H1000&) Or ((thirdByte And &H3F&) * &H40&) Or (fourthByte And &H3F&)
            inputPosition = inputPosition + 4
        Else
            Err.Raise 5, "Utf8.DecodeUTF8", "Invalid UTF-8 sequence."
        End If

        If codepoint <= &HFFFF& Then
            Mid$(outputText, outputPosition, 1) = UnicodeCharacter(codepoint)
            outputPosition = outputPosition + 1
        Else
            codepoint = codepoint - &H10000
            Mid$(outputText, outputPosition, 1) = UnicodeCharacter(&HD800& + (codepoint \ &H400&))
            Mid$(outputText, outputPosition + 1, 1) = UnicodeCharacter(&HDC00& + (codepoint And &H3FF&))
            outputPosition = outputPosition + 2
        End If
    Loop

    DecodeUTF8 = Left$(outputText, outputPosition - 1)
End Function

Private Function UnicodeCharacter(ByVal codeUnit As Long) As String
    If codeUnit > &H7FFF& Then codeUnit = codeUnit - &H10000
    UnicodeCharacter = ChrW$(codeUnit)
End Function

Private Sub EncodeUTF8(ByRef utf16leStr As String, ByRef utf8() As Byte, _
              Optional ByVal raiseErrors As Boolean = False)
    Const methodName As String = "EncodeUTF8"
    Dim codepoint As Long
    Dim lowSurrogate As Long
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0
    ReDim utf8(Len(utf16leStr) * 4 - 1)

    Do While i <= Len(utf16leStr)
        codepoint = AscW(Mid$(utf16leStr, i, 1)) And &HFFFF&

        If codepoint >= &HD800& And codepoint <= &HDBFF& Then 'high surrogate
            If i < Len(utf16leStr) Then
                lowSurrogate = AscW(Mid$(utf16leStr, i + 1, 1)) And &HFFFF&
            Else
                lowSurrogate = 0 'Will fail the validity check below
            End If

            If &HDC00& <= lowSurrogate And lowSurrogate <= &HDFFF& Then
                codepoint = (codepoint - &HD800&) * &H400& + _
                            (lowSurrogate - &HDC00&) + &H10000
                i = i + 1
            Else
                If raiseErrors Then _
                    Err.Raise 5, methodName, _
                        "Invalid Unicode codepoint. (Lonely high surrogate)"
                codepoint = &HFFFD&
            End If
        End If

        If codepoint < &H80& Then
            utf8(j) = codepoint
            j = j + 1
        ElseIf codepoint < &H800& Then
            utf8(j) = &HC0& Or ((codepoint And &H7C0&) \ &H40&)
            utf8(j + 1) = &H80& Or (codepoint And &H3F&)
            j = j + 2
        ElseIf codepoint < &HDC00& Then
            utf8(j) = &HE0& Or ((codepoint And &HF000&) \ &H1000&)
            utf8(j + 1) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 2) = &H80& Or (codepoint And &H3F&)
            j = j + 3
        ElseIf codepoint < &HE000& Then
            If raiseErrors Then _
                Err.Raise 5, methodName, _
                    "Invalid Unicode codepoint. (Lonely low surrogate)"
            codepoint = &HFFFD&
            utf8(j) = &HE0& Or ((codepoint And &HF000&) \ &H1000&)
            utf8(j + 1) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 2) = &H80& Or (codepoint And &H3F&)
            j = j + 3
        ElseIf codepoint < &H10000 Then
            utf8(j) = &HE0& Or ((codepoint And &HF000&) \ &H1000&)
            utf8(j + 1) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 2) = &H80& Or (codepoint And &H3F&)
            j = j + 3
        Else
            utf8(j) = &HF0& Or ((codepoint And &H1C0000) \ &H40000)
            utf8(j + 1) = &H80& Or ((codepoint And &H3F000) \ &H1000&)
            utf8(j + 2) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 3) = &H80& Or (codepoint And &H3F&)
            j = j + 4
        End If

        i = i + 1
    Loop
    ReDim Preserve utf8(0 To j - 1)
End Sub
