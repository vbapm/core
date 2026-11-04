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
