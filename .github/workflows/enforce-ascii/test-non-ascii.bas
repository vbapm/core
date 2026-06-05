VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Test file with non-ASCII character: é
' This should be caught by enforce-ascii workflow
Public Sub Hello()
    MsgBox "Héllo"
End Sub

Public Function GetName() As String
    GetName = "José"
End Function
