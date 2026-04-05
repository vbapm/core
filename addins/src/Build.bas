Attribute VB_Name = "Build"
''
' # Build
'
' Primary build tooling for import, export, and create
'
' @module Build
' @author Tim Hall <tim.hall.engr@gmail.com>
' @repository https://github.com/vba-blocks/vba-blocks
' @license MIT
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function ImportGraph(Graph As Variant) As String
    On Error GoTo ErrorHandling

    Dim Values As Dictionary
    Dim Document As Object
    Dim App As New OfficeApplication

    Set Values = JsonConverter.ParseJson(Graph)
    Set Document = App.GetDocument(Values("file"))

    DebugLog.Clear
    DebugLog.Log "ImportGraph", "Starting import for: " & Values("file")

    ' Capture the currently active VBE code pane (best-effort, non-fatal)
    Dim ActiveComponentName As String
    Dim ActiveStartLine As Long, ActiveStartCol As Long
    Dim ActiveEndLine As Long, ActiveEndCol As Long
    CaptureActiveCodePane Document.VBProject, ActiveComponentName, ActiveStartLine, ActiveStartCol, ActiveEndLine, ActiveEndCol

    Dim Src As Dictionary
    For Each Src In Values("src")
        Output.Messages.Add "src: " & Src("name") & ", " & Src("path")
        Installer.Import Document.VBProject, Src("name"), Src("path"), Overwrite:=True
    Next Src

    Dim Ref As Dictionary
    For Each Ref In Values("references")
        Output.Messages.Add "ref: " & Ref("name") & ", " & Ref("guid") & ", " & Ref("major") & ", " & Ref("minor")
        Installer.AddReference Document.VBProject, Ref("guid"), CLng(Ref("major")), CLng(Ref("minor"))
    Next Ref

    Document.Save

    ' Restore the previously active code pane (best-effort, non-fatal)
    DebugLog.Log "ImportGraph", "Restore: ComponentName='" & ActiveComponentName & "' StartLine=" & ActiveStartLine & " StartCol=" & ActiveStartCol
    RestoreActiveCodePane Document.VBProject, ActiveComponentName, ActiveStartLine, ActiveStartCol, ActiveEndLine, ActiveEndCol

    ImportGraph = Output.Result
    Exit Function

ErrorHandling:

    Output.Errors.Add Err.Number & ": " & Err.Description
    ImportGraph = Output.Result
End Function

''
' Export given file to the given staging directory
'
' @param {String} Info json value for file and staging
' @param {String} Info.file absolute file path to document to export
' @param {String} Info.staging absolute path to "staging" directory to export to
''
Public Function ExportTo(Info As Variant) As String
    On Error GoTo ErrorHandling

    Dim Values As Dictionary
    Dim Staging As String
    Dim Document As Object
    Dim App As New OfficeApplication

    Output.Messages.Add "ExportTo"

    Set Values = JsonConverter.ParseJson(Info)
    Set Document = App.GetDocument(Values("file"))
    Staging = Values("staging")

    ' Iterate through all components in document and export directly to staging
    Dim Component As VBComponent
    Dim Path As String
    For Each Component In Document.VBProject.VBComponents
        Dim Extension As String
        Select Case Component.Type
        Case vbext_ComponentType.vbext_ct_StdModule
            Extension = ".bas"
        Case vbext_ComponentType.vbext_ct_ClassModule, vbext_ComponentType.vbext_ct_Document
            Extension = ".cls"
        Case vbext_ComponentType.vbext_ct_MSForm
            Extension = ".frm"
        Case Else
            ' The only other component type for Excel is vbext_ct_ActiveXDesigner = 11
            ' I'm not sure when this could occur, so just warn for now
            Output.Warnings.Add "Unknown component type: " & Component.Type
        End Select

        ' Avoid exporting built-in modules / classes that are blank
        ' User-added modules / classes that are blank are assumed to be intentional
        If Extension <> "" And Not (Component.Type = vbext_ComponentType.vbext_ct_Document And ComponentIsBlank(Component)) Then
            Path = FileSystem.JoinPath(Staging, Component.Name & Extension)
            Installer.Export Document.VBProject, Component.Name, Path, Overwrite:=True
        End If
    Next Component

    ' For "indirect" values (VBA project name and references)
    ' export to project.json for post-processing by vbapm
    Dim Project As New Dictionary

    Project("name") = Document.VBProject.Name
    Set Project("references") = New Collection

    Dim Ref As Reference
    Dim RefInfo As Dictionary
    Dim BuiltInReferences As Variant
    BuiltInReferences = Array("stdole", "office", "msforms")
    For Each Ref In Document.VBProject.References
        If Not Ref.BuiltIn And Not InArray(VBA.LCase$(Ref.Name), BuiltInReferences) Then
            Set RefInfo = New Dictionary
            RefInfo("name") = Ref.Name
            RefInfo("version") = Ref.Major & "." & Ref.Minor
            RefInfo("guid") = Ref.Guid
            RefInfo("major") = Ref.Major
            RefInfo("minor") = Ref.Minor

            Project("references").Add RefInfo
        End If
    Next Ref

    Dim ProjectPath As String
    Dim ProjectJson As String

    ProjectPath = FileSystem.JoinPath(Staging, "project.json")
    ProjectJson = JsonConverter.ConvertToJson(Project)

    Open ProjectPath For Output As #1
    Print #1, ProjectJson
    Close #1

    ExportTo = Output.Result
    Exit Function

ErrorHandling:

    Output.Errors.Add Err.Number & ": " & Err.Description
    ExportTo = Output.Result
End Function

''
' Create a blank document at path
'
' @param {String} Info.path
''
Public Function CreateDocument(Info As Variant) As String
    On Error GoTo ErrorHandling

    Dim Values As Dictionary
    Dim DocumentPath As String
    Dim App As New OfficeApplication

    Set Values = JsonConverter.ParseJson(Info)
    App.CreateDocument Values("path")

    CreateDocument = Output.Result
    Exit Function

ErrorHandling:

    Output.Errors.Add Err.Number & ": " & Err.Description
    CreateDocument = Output.Result
End Function

' ============================================= '

''
' Capture the active VBE CodePane if it belongs to the given VBProject.
' Stores the component name and cursor selection for later restoration.
' All errors are silently ignored - this is a best-effort, non-fatal operation.
''
Private Sub CaptureActiveCodePane( _
    Project As VBProject, _
    ByRef ComponentName As String, _
    ByRef StartLine As Long, _
    ByRef StartCol As Long, _
    ByRef EndLine As Long, _
    ByRef EndCol As Long _
)
    'On Error Resume Next

    DebugLog.Log "CaptureActiveCodePane", "Checking VBE.ActiveCodePane"

    Dim ActivePane As Object
    Set ActivePane = Application.VBE.ActiveCodePane
    If ActivePane Is Nothing Then
        DebugLog.Log "CaptureActiveCodePane", "No active code pane ? skipping"
        Exit Sub
    End If

    DebugLog.Log "CaptureActiveCodePane", "Active pane component: '" & ActivePane.CodeModule.Parent.Name & "'"

    ' Only capture if the pane belongs to the workbook being updated
    If Not ActivePane.CodeModule.Parent.Collection.Parent Is Project Then
        DebugLog.Log "CaptureActiveCodePane", "Pane belongs to a different project ? skipping"
        Exit Sub
    End If

    ComponentName = ActivePane.CodeModule.Parent.Name
    ActivePane.GetSelection StartLine, StartCol, EndLine, EndCol
    DebugLog.Log "CaptureActiveCodePane", "Captured: '" & ComponentName & "' L" & StartLine & ":C" & StartCol & " to L" & EndLine & ":C" & EndCol
End Sub

''
' Restore the previously captured VBE CodePane in the given VBProject.
' All errors are silently ignored - this is a best-effort, non-fatal operation.
''
Private Sub RestoreActiveCodePane( _
    Project As VBProject, _
    ComponentName As String, _
    StartLine As Long, _
    StartCol As Long, _
    EndLine As Long, _
    EndCol As Long _
)
    'On Error Resume Next

    If ComponentName = "" Then
        DebugLog.Log "RestoreActiveCodePane", "No component name captured ? nothing to restore"
        Exit Sub
    End If

    DebugLog.Log "RestoreActiveCodePane", "Looking for component: '" & ComponentName & "'"

    ' The component was just removed and re-imported, so its old CodePane is gone.
    ' Look it up by name in VBComponents and access CodeModule.CodePane ? this
    ' property opens the pane automatically if it is not already open.
    Dim Component As VBComponent
    Set Component = Project.VBComponents(ComponentName)

    DebugLog.Log "RestoreActiveCodePane", "Found component, accessing CodePane"

    Dim cp As Object
    Set cp = Component.CodeModule.CodePane

    DebugLog.Log "RestoreActiveCodePane", "Activating and setting selection L" & StartLine & ":C" & StartCol
    Set Application.VBE.ActiveCodePane = cp
    cp.SetSelection StartLine, StartCol, EndLine, EndCol
End Sub

Private Function ComponentIsBlank(Component As VBComponent) As Boolean
    Dim LineNumber As Long
    Dim Line As String

    For LineNumber = 1 To Component.CodeModule.CountOfLines
        Line = Component.CodeModule.Lines(LineNumber, 1)
        If Not (Line = "Option Explicit" Or Line = "") Then
            ComponentIsBlank = False
            Exit Function
        End If
    Next LineNumber

    ComponentIsBlank = True
End Function

Private Function InArray(Value As Variant, Values As Variant) As Boolean
    Dim i As Long
    For i = LBound(Values) To UBound(Values)
        If Values(i) = Value Then
            InArray = True
            Exit Function
        End If
    Next i
End Function
