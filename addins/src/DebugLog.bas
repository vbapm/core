Attribute VB_Name = "DebugLog"
''
' # DebugLog
'
' Simple file-based logger for debugging. Writes timestamped lines to
' %TEMP%\vbapm-debug.log (Windows) or /tmp/vbapm-debug.log (Mac).
'
' Usage:
'   DebugLog.Log "MyModule", "something happened: " & someValue
'   DebugLog.Clear   ' optional: wipe the log at the start of an operation
'
' @module DebugLog
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Function LogPath() As String
    Dim tmp As String

#If Mac Then
    tmp = "/tmp"
#Else
    tmp = Environ("TEMP")
    If tmp = "" Then tmp = Environ("TMP")
    If tmp = "" Then tmp = "C:\Temp"
#End If

    LogPath = tmp & Application.PathSeparator & "vbapm-debug.log"
End Function

''
' Append a timestamped line to the log file.
' @param {String} Source   Module or procedure name
' @param {String} Message  The message to log
''
Public Sub Log(Source As String, Message As String)
    Dim f As Integer
    f = FreeFile

    Open LogPath() For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:mm:ss") & " [" & Source & "] " & Message
    Close #f
End Sub

''
' Clear the log file (call once at the start of an operation).
''
Public Sub Clear()
    Dim f As Integer
    f = FreeFile

    Open LogPath() For Output As #f
    Close #f
End Sub
