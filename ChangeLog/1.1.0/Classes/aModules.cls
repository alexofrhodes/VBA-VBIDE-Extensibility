VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun AddOrSet       As aModule     :
' Fun ClassNames     As Variant     :
' Fun Classes        As Collection  :
' Fun Documents      As Collection  :
' Fun Exists         As Boolean     :
' Fun Initialize     As aModules    :
' Fun Item           As aModule     :
' Fun ModuleNames    As Variant     :
' Fun NamesOf        As Variant     :
' Fun NormalModules  As Collection  :
' Fun UserformNames  As Variant     :
' Fun Userforms      As Collection  :
' Get Items          As Variant     :
' Sub CaseProperModulesOfWorkbook   :
' Sub CommentsRemove                :
' Sub Export                        :
' Sub ExportProcedures              :
' Sub ImportPaths                   :
' Sub ImportPicker                  :
' Sub Indent                        :
' Sub InjectOptionExplicit          :
' Sub ListProcedures                :
' Sub PrintTodoList                 :
' Sub Refresh                       :
' Sub RemoveEmptyLinesButLeaveOne   :
' Sub RemoveLinesLike               :
' Sub RemoveProcedureList           :
' Sub SideBySide                    :
' Sub UpdateProcedures              :
'--------------------------------------------------
'@EndFolder ListOfProcedures

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aModules
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 29-06-2023 13:46    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private oProject As VBProject

Public Function Initialize(TargetProject As VBProject) As aModules
    Set oProject = TargetProject
    Set Initialize = Me
End Function

Public Property Get Items(Optional index)
    Dim am As aModule
    If IsMissing(index) Then
        Set Items = New Collection
        Dim TargetModule As VBComponent
        For Each TargetModule In oProject.VBComponents
            Set am = New aModule: am.Initialize TargetModule
            Items.Add am
        Next
    ElseIf TypeName(index) = "String" Or IsNumeric(index) Then
        Set am = New aModule
        am.Initialize oProject.VBComponents(index)
        Set Items = am
    End If
End Property

Function NormalModules() As Collection
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim coll        As New Collection
    Dim element     As aModule
    For Each element In Items
        If element.TypeToLong = vbext_ct_StdModule Then
            coll.Add element
        End If
    Next
    Set NormalModules = coll
End Function

Function Userforms() As Collection
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim coll        As New Collection
    Dim element     As aModule
    For Each element In Items
        If element.TypeToLong = vbext_ct_MSForm Then
            coll.Add element
        End If
    Next
    Set Userforms = coll
End Function

Function Classes() As Collection
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim coll        As New Collection
    Dim element     As aModule
    For Each element In Items
        If element.TypeToLong = vbext_ct_ClassModule Then
            coll.Add element
        End If
    Next
    Set Classes = coll
End Function

Function Documents() As Collection
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim coll        As New Collection
    Dim element     As aModule
    For Each element In Items
        If element.TypeToLong = vbext_ct_Document Then
            coll.Add element
        End If
    Next
    Set Documents = coll
End Function

Sub Export(targetPath As String, createSubfolders As Boolean)
    '@AssignedModule aModules
    '@INCLUDE PROCEDURE FoldersCreate
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim ClassFolder As String: ClassFolder = targetPath
    Dim FormFolder  As String: FormFolder = targetPath
    Dim NormalModuleFolder As String: NormalModuleFolder = targetPath
    Dim DocumentFolder As String: DocumentFolder = targetPath

    If createSubfolders Then
        FormFolder = targetPath & "Userforms"
        NormalModuleFolder = targetPath & "Modules"
        DocumentFolder = targetPath & "Documents"
        ClassFolder = targetPath & "Classes"

        FoldersCreate FormFolder
        FoldersCreate NormalModuleFolder
        FoldersCreate ClassFolder
        FoldersCreate DocumentFolder
    End If

    Dim am As aModule
    For Each am In Me.Classes:          am.Export ClassFolder: Next am
    For Each am In Me.Userforms:        am.Export FormFolder: Next am
    For Each am In Me.NormalModules:    am.Export NormalModuleFolder: Next am
    For Each am In Me.Documents:        am.Export DocumentFolder: Next am
End Sub



Sub PrintTodoList()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim am          As aModule
    For Each am In Items
        am.PrintTodoList
    Next
End Sub


Sub InjectOptionExplicit()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim am          As aModule
    For Each am In Items
        If Not am.Contains("Option Explicit", False, True, False) Then am.Component.CodeModule.InsertLines 1, "Option Explicit"
    Next
End Sub



Public Sub Indent()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim Module      As aModule
    For Each Module In Me.Items
        Module.Indent
    Next
End Sub
Sub RemoveEmptyLinesButLeaveOne()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim Module      As aModule
    For Each Module In Items
        Module.RemoveEmptyLinesButLeaveOne
    Next
End Sub
Sub UpdateProcedures()
    '@AssignedModule aModules
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim Procedure   As aProcedure
    Dim Module      As aModule
    For Each Module In Items
        If Module.TypeToLong = vbext_ct_StdModule Then
            For Each Procedure In Module.Procedures
                If UCase(CStr(Procedure)) <> UCase("UpdateAllProcedures") Then
                    Procedure.Dependencies.Update
                End If
            Next
        End If
    Next
End Sub



Public Function Item(IndexOrName As Variant) As aModule
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Set Item = Items.Item(IndexOrName)
End Function

Public Function Exists(TargetName As String) As Boolean
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim Module      As VBComponent
    On Error Resume Next
    Set Module = oProject.VBComponents(TargetName)
    On Error GoTo 0
    Exists = Not Module Is Nothing
End Function

Function AddOrSet( _
        TargetName As String, _
        moduleType As VBIDE.vbext_ComponentType) As aModule
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules

    Dim Module      As VBComponent
    On Error Resume Next
    Set Module = oProject.VBComponents(TargetName)
    On Error GoTo 0
    If Module Is Nothing Then
        Set Module = oProject.VBComponents.Add(moduleType)
        Module.Name = TargetName
    End If
    Dim am          As New aModule
    am.Initialize Module
    Set AddOrSet = am
End Function

Sub ImportPaths(ReplaceExisting As Boolean, ParamArray Paths() As Variant)
    '@AssignedModule aModules
    '@INCLUDE PROCEDURE Toast
    '@INCLUDE PROCEDURE DataFilePartExtension
    '@INCLUDE PROCEDURE DataFilePartName
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    If LBound(Paths) = -1 Then Exit Sub

    Dim Extension   As String
    Dim TargetName  As String
    Dim element

    Dim PASS        As Boolean
    For Each element In Paths
        Select Case True
            Case element Like "*.cls", "*.frm", "*.bas", "*.bas"
            Case Else: Exit Sub
        End Select
    Next
    For Each element In Paths
        TargetName = DataFilePartName(CStr(element), False)
        Extension = DataFilePartExtension(CStr(element))
        If Exists(TargetName) Then
            If ReplaceExisting Then
                aModule.Initialize(oProject.VBComponents(TargetName)).Delete
                oProject.VBComponents.Import CStr(element)
            End If
        Else
            oProject.VBComponents.Import CStr(element)
        End If
    Next element

    'https://github.com/rfl808/Notify
    Toast , "Import successful"
End Sub

Sub ImportPicker(ReplaceExisting As Boolean)
    '@AssignedModule aModules
    '@INCLUDE PROCEDURE ArrayAllocated
    '@INCLUDE PROCEDURE WorksheetExists
    '@INCLUDE PROCEDURE Toast
    '@INCLUDE PROCEDURE DataFilePartFolder
    '@INCLUDE PROCEDURE DataFilePicker
    '@INCLUDE PROCEDURE DataFilePartExtension
    '@INCLUDE PROCEDURE DataFilePartName
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim SelectedModules: SelectedModules = DataFilePicker(Array("bas", "frm", "cls"), True)
    If Not ArrayAllocated(SelectedModules) Then Exit Sub
    Dim basePath    As String: basePath = DataFilePartFolder(SelectedModules(1), True)
    '    Dim SourceWorkbook     As Workbook
    '    Dim SourceWorkbookName As String:   SourceWorkbookName = Dir(basePath & "*.xl*")
    '    Dim wasOpen            As Boolean
    '
    '    If SourceWorkbookName <> "" Then: wasOpen = WorkbookExists(SourceWorkbookName)

    Dim Extension   As String
    Dim TargetName  As String
    Dim element
    For Each element In SelectedModules
        TargetName = DataFilePartName(CStr(element), False)
        Extension = DataFilePartExtension(CStr(element))
        '        If UCase(Extension) <> UCase("doccls") Then
        If Exists(TargetName) Then
            If ReplaceExisting Then
                aModule.Initialize(oProject.VBComponents(TargetName)).Delete
                oProject.VBComponents.Import CStr(element)
            End If
        Else
            oProject.VBComponents.Import CStr(element)
        End If

        '        ElseIf UCase(Extension) = UCase("doccls") And SourceWorkbookName <> "" Then
        '            If WorksheetExists(TargetName, oWorkbook) Then
        '                If ReplaceExisting Then
        ''                    TargetWorkbook.Worksheets.Copy
        '                    TargetWorkbook.Worksheets(TargetName).Delete
        '                Else
        '                    GoTo NextElement
        '                End If
        '
        '                If wasOpen = False Then
        '                    Application.EnableEvents = False
        '                    Set SourceWorkbook = Workbooks.Open(basePath & SourceWorkbookName)
        '                Else
        '                    Set SourceWorkbook = Workbooks(SourceWorkbookName)
        '                End If
        '                SourceWorkbook.SHEETS(TargetName).Copy Before:=TargetWorkbook.SHEETS(1)
        '                Application.EnableEvents = True
        '            End If
        '        End If
NextElement:
    Next element

    '    If wasOpen = False And WorkbookExists(SourceWorkbookName) Then SourceWorkbook.Close False
    'https://github.com/rfl808/Notify
    Toast , "Import successful"
End Sub

Sub ListProcedures()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim Module      As VBComponent
    For Each Module In oProject.VBComponents
        If Module.Type <> vbext_ct_ClassModule And Module.Type <> vbext_ct_Document Then
            If Module.CodeModule.countOfLines > 0 Then
                aModule.Initialize(Module).Procedures.List
            End If
        End If
    Next
End Sub

Sub ListProceduresPublic()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim Module      As VBComponent
    For Each Module In oProject.VBComponents
        If Module.Type <> vbext_ct_ClassModule And Module.Type <> vbext_ct_Document Then
            If Module.CodeModule.countOfLines > 0 Then
                aModule.Initialize(Module).Procedures.ListPublic
            End If
        End If
    Next
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 13:55    Alex                (aModules.cls > RemoveLinesLike)

Public Sub RemoveLinesLike(this As String)
'@LastModified 2308221355
    Dim Module      As aModule
    For Each Module In Items
        Module.RemoveLinesLike this
    Next
End Sub

'@TODO
Sub SideBySide(ParamArray TwoModuleNamesOrOneFormName() As Variant)
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules

    Dim Module1     As VBComponent
    Dim Module2     As VBComponent
    Set Module1 = oProject.VBComponents(TwoModuleNamesOrOneFormName(0))
    If UBound(TwoModuleNamesOrOneFormName) > 0 Then Set Module2 = oProject.VBComponents(TwoModuleNamesOrOneFormName(1))
    With Module1.CodeModule.CodePane.Window
        .Width = 800
        .Left = 1
        .Top = 1
        .Height = 932
        .Visible = True
        .WindowState = vbext_ws_Normal
        .SetFocus
    End With
    If Not Module2 Is Nothing Then
        With Module2.CodeModule.CodePane.Window
            .Width = 800
            .Left = 800
            .Top = 1
            .Height = 932
            .Visible = True
            .WindowState = vbext_ws_Normal
            .SetFocus
        End With
    ElseIf Module1.Type = vbext_ct_MSForm Then
        With Module1.DesignerWindow
            .Width = 800
            .Left = 800
            .Top = 1
            .Height = 932
            .Visible = True
            .WindowState = vbext_ws_Normal
            Module1.DesignerWindow.SetFocus
        End With
        Exit Sub

    End If
End Sub
'
'Sub CaseProperModulesOfWorkbook(Optional TargetWorkbook As Workbook)
'    Dim Module As VBComponent
'    For Each Module In TargetWorkbook.VBProject.VBComponents
'        If Module.Name <> "ThisWorkbook" Then
'            Module.Name = UCase(Left(Module.Name, 1)) & Mid(Module.Name, 2)
'        End If
'    Next
'End Sub

Public Sub Refresh()
    '@AssignedModule aModules
    '@INCLUDE PROCEDURE Toast
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    If oProject.fileName <> ThisWorkbook.FullName Then
        ExportModules
        ImportModules
    Else
        Toast "Can't touch this"
    End If
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 30-10-2023 09:43    Alex                (aModules.cls > ExportProcedures : + argument to create subfolders by module name)

Sub ExportProcedures(targetPath As String, Optional createModuleSubfolders As Boolean)
'@LastModified 2310300943
    '@AssignedModule aModules
    '@INCLUDE PROCEDURE FoldersCreate
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    targetPath = Replace(targetPath & "\", "\\", "\")
    Dim ProcedurePath As String
    Dim am          As aModule
    Dim ap          As aProcedure
    Dim coll        As New Collection
    Dim ModulePath  As String
    For Each am In NormalModules 'Userforms and Classes should be exported as they are
        If createModuleSubfolders Then
            ModulePath = targetPath & am.Name & "\"
            FoldersCreate ModulePath
            am.Procedures.Export ModulePath
        Else
            am.Procedures.Export targetPath
        End If
    Next
End Sub

Private Sub ExportModules()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim bExport     As Boolean
    Dim szExportPath As String
    Dim szFileName  As String
    Dim cmpComponent As VBIDE.VBComponent
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    On Error Resume Next
    Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0
    szExportPath = FolderWithVBAProjectFiles & "\"
    For Each cmpComponent In oProject.VBComponents
        bExport = True
        szFileName = cmpComponent.Name
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                bExport = False
        End Select
        If bExport Then
            cmpComponent.Export szExportPath & szFileName
        End If
    Next cmpComponent
End Sub

Private Sub ImportModules()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim objFSO      As Scripting.FileSystemObject
    Dim objFile     As Scripting.File
    Dim szImportPath As String
    Dim szFileName  As String
    Dim cmpComponents As VBIDE.VBComponents
    If oProject.fileName = ThisWorkbook.FullName Then
        MsgBox "Select another destination workbook" & _
                "Not possible to import in this workbook "
        Exit Sub
    End If
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If
    szImportPath = FolderWithVBAProjectFiles & "\"
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
        MsgBox "There are no files to import"
        Exit Sub
    End If
    Call DeleteVBAModulesAndUserForms
    Set cmpComponents = oProject.VBComponents
    For Each objFile In objFSO.GetFolder(szImportPath).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.path
        End If
    Next objFile
End Sub

Private Function FolderWithVBAProjectFiles() As String
    '@AssignedModule aModules
    '@INCLUDE PROCEDURE FolderExists
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim WshShell    As Object
    Dim FSO         As Object
    Dim SpecialPath As String
    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")
    SpecialPath = WshShell.SpecialFolders("MyDocuments")
    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
End Function

Private Sub DeleteVBAModulesAndUserForms()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim vbComp      As VBIDE.VBComponent
    For Each vbComp In oProject.VBComponents
        If vbComp.Type = vbext_ct_Document Then
        Else
            oProject.VBComponents.Remove vbComp
        End If
    Next vbComp
End Sub

Public Sub CommentsRemove()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim am          As aModule
    For Each am In Items
        am.CommentsRemove
    Next
End Sub

Sub RemoveProcedureList()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim am          As aModule
    For Each am In Items
        am.Folders.Delete "ListOfProcedures"
    Next
End Sub

Function ClassNames()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Set ClassNames = NamesOf(vbext_ct_ClassModule)
End Function
Function ModuleNames()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Set ModuleNames = NamesOf(vbext_ct_StdModule)
End Function

Function UserformNames()
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Set UserformNames = NamesOf(vbext_ct_MSForm)
End Function

Sub CaseProperModulesOfWorkbook()
    Dim Module As VBComponent
    For Each Module In oProject.VBComponents
        If Module.Name <> "ThisWorkbook" Then
            Module.Name = UCase(Left(Module.Name, 1)) & Mid(Module.Name, 2)
        End If
    Next
End Sub


Public Function NamesOf(moduleType As vbext_ComponentType)
    '@AssignedModule aModules
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim coll        As New Collection
    Dim Module      As VBComponent
    For Each Module In oProject.VBComponents
        If Module.Type = moduleType Then
            coll.Add Module.Name
        End If
    Next
    Set NamesOf = coll
End Function


