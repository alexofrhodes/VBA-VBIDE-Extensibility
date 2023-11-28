VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Code             As String                :
' Fun Declarations     As aProjectDeclarations  :
' Fun Extension        As Variant               :
' Fun Name             As Variant               :
' Fun NameClean        As Variant               :
' Fun ProceduresArray  As Variant               :
' Fun ProceduresLike   As Collection            :
' Fun REFERENCES       As aProjectReferences    :
' Fun TodoList         As String                :
' Fun WorkbookObject   As Workbook              :
' Get Active           As aProject              :
' Get Initialize       As aProject              :
' Get Items            As Variant               :
' Get Procedures       As Collection            :
' Get ProceduresNames  As Collection            :
' Get Project          As VBProject             :
' Get this             As aProject              :
' Sub Backup                                    :
' Sub CreateLinkedTable                         :
' Sub CreateLinkedTableSheet                    :
' Sub ExportCodeUnified                         :
' Sub ExportModules                             :
' Sub ExportProcedures                          :
' Sub ExportXML                                 :
' Sub Indent                                    :
' Sub ModulesMerge                              :
'--------------------------------------------------
'@EndFolder ListOfProcedures
Option Explicit

Private oProject As VBProject

Private Type tAddons
    REFERENCES As aProjectReferences
    Declarations As aProjectDeclarations
End Type

Private Addons As tAddons

Public Property Get Initialize(index) As aProject
'eg we have in VBA Project Explorer :  ProjectName(Filename.xlsm)
'initialize 1
'initialize "ProjectName"
'initialize "Filename.xlsm"
'initialize workbooks("filename.xlsm")
'initialize thisworkbook.vbProject
    Set oProject = Nothing
    If TypeName(index) = "VBProject" Then Set oProject = index
    If oProject Is Nothing Then
        On Error Resume Next
        Set oProject = Application.VBE.VBProjects(index)
        On Error GoTo 0
    End If
    If oProject Is Nothing Then
        On Error Resume Next
        Set oProject = index.VBProject
        On Error GoTo 0
    End If
    If oProject Is Nothing Then
        index = Split(index, "\")
        index = index(UBound(index))
        Dim TargetProject As VBProject
        Dim TargetProjectName
        For Each TargetProject In Application.VBE.VBProjects
            TargetProjectName = Split(TargetProject.fileName, "\")
            TargetProjectName = TargetProjectName(UBound(TargetProjectName))
            If UCase(TargetProjectName) = UCase(index) Then
                Set oProject = TargetProject
                Exit For
            End If
        Next
    End If
    If oProject Is Nothing Then GoTo ErrorHandler
    Set Initialize = Me
    Exit Property
ErrorHandler:

End Property

Public Property Get Active() As aProject
    Dim Module As VBComponent
    Dim am As aModule: Set am = New aModule
    Set Module = am.Active.Component
    If Not Module Is Nothing Then
        Set Active = Initialize(Module.Collection.Parent)
    Else
    
    End If
End Property

Public Property Get this() As aProject
'@TODO
'    Dim appName As String
'    appName = Application.Name
'    #If Application.Name = "Microsoft Excel" Then
        Set this = Initialize(ThisWorkbook.VBProject.fileName)
'    #ElseIf Application.Name = "Microsoft Word" Then
'        Set This = ThisDocument.VBProject
'    #ElseIf Application.Name = "Microsoft PowerPoint" Then
'        '// TODO
'    #ElseIf Application.Name = "Microsoft Access" Then
'        '// TODO
'    #ElseIf Application.Name = "Microsoft Outlook" Then
'        '// TODO
'    #End If
End Property

Public Property Get Items(Optional index As Variant)
    Dim ams As aModules: Set ams = New aModules
    ams.Initialize oProject
    If IsMissing(index) Then
        Set Items = ams.Items
    Else
        Set Items = ams.Items(index)
    End If
End Property

Public Function REFERENCES() As aProjectReferences
    If Addons.REFERENCES Is Nothing Then
        Set Addons.REFERENCES = New aProjectReferences
        Addons.REFERENCES.Initialize Me
    End If
    Set REFERENCES = Addons.REFERENCES
End Function

Public Function Declarations() As aProjectDeclarations
    If Addons.Declarations Is Nothing Then
        Set Addons.Declarations = New aProjectDeclarations
        Addons.Declarations.Initialize Me
    End If
    Set Declarations = Addons.Declarations
End Function

Public Property Get Project() As VBProject
    Set Project = oProject
End Property

Public Function WorkbookObject() As Workbook
    Dim bookName As Variant
    bookName = Split(oProject.fileName, "\")
    bookName = bookName(UBound(bookName))
    Set WorkbookObject = Workbooks(bookName)
End Function


Function Name()
    Dim output
    output = Split(oProject.fileName, "\")
    output = output(UBound(output))
    Name = output
End Function
Function Extension()
    Extension = Mid(Name, InStr(1, Name, "."))
End Function

Function NameClean()
    NameClean = Mid(Name, 1, InStr(1, Name, ".") - 1)
End Function

Public Sub Indent()
    aModules.Initialize(oProject).Indent
End Sub

Sub Backup(Optional targetFolder As String)
    If targetFolder = "" Then targetFolder = WorkbookObject.path & "\Backup\"
    targetFolder = VBA.Replace(targetFolder & "\", "\\", "\")
    FoldersCreate targetFolder
    WorkbookObject.SaveCopyAs targetFolder & _
            Format(Now, "yyyy-mm-dd hh-nn") & " " & Name
End Sub

Public Function code() As String
    Dim am      As aModule
    Dim txt
    For Each am In aModules.Initialize(oProject).Items
        If am.Component.CodeModule.countOfLines > 0 Then
            txt = txt & _
                    vbNewLine & _
                    "'" & String(10, "=") & " (" & am.TypeToString & ") " & am.Name & String(10, "=") & _
                    vbNewLine & _
                    am.code
        End If
    Next
    code = txt
End Function


Function TodoList() As String
    Dim output         As String
    Dim am          As aModule
    For Each am In aModules.Initialize(oProject).Items
        If InStr(1, am.code, "@TODO", vbTextCompare) > 0 Then
            output = output & IIf(output <> "", vbNewLine, "") & am.TodoList
        End If
    Next am
    TodoList = output
End Function

Public Property Get Procedures( _
                            includeModule As Boolean, _
                            includeDocument As Boolean, _
                            includeClass As Boolean, _
                            includeForm As Boolean) As Collection
    'as aProcedure objects
    Dim Module      As VBComponent
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim ProcedureName As String
    Dim cProcedure  As aProcedure
    Dim i           As Long
    For Each Module In oProject.VBComponents
        If Not includeModule And Module.Type = vbext_ct_StdModule Then GoTo SKIP
        If Not includeClass And Module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If Not includeDocument And Module.Type = vbext_ct_Document Then GoTo SKIP
        If Not includeForm And Module.Type = vbext_ct_MSForm Then GoTo SKIP
        With Module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .countOfLines

                For i = 0 To 4
                    procKind = i
                    ProcedureName = .ProcOfLine(lineNum, procKind)
                    If ProcedureName <> vbNullString Then

                        'Possible Error:
                        '---------------
                        ' Procedures with an underscore _ which is used in events
                        ' and Private Procedures
                        ' may have the same name in different components

                        If InStr(1, ProcedureName, "_") = 0 Then
                            ProcedureName = .ProcOfLine(lineNum, procKind)
                            Set cProcedure = New aProcedure
                            cProcedure.Initialize WorkbookObject, Module, ProcedureName, procKind
                            On Error Resume Next
                            coll.Add cProcedure, cProcedure.Name & procKind
                            'coll.Add ProcedureName, ProcedureName
                            'If Err.Number > 0 Then Debug.Print ProcedureName & " exists multiple times, added only once, may lead to error"
                            'Err.clear
                            On Error GoTo 0
                        End If

                        Exit For

                    End If
                Next
                lineNum = .procStartLine(ProcedureName, procKind) + .ProcCountLines(ProcedureName, procKind) + 1
            Loop
        End With
SKIP:
    Next Module

    Set Procedures = coll
End Property

Function ProceduresLike(this As String) As Collection
    Dim Procedure   As aProcedure
    Dim out         As New Collection
    For Each Procedure In Procedures(True, False, False, False)
        If UCase(Procedure.Name) Like UCase(this) Then
            out.Add Procedure
        End If
    Next
    Set ProceduresLike = out
End Function

Public Property Get ProceduresNames( _
        includeModule As Boolean, _
        includeDocument As Boolean, _
        includeClass As Boolean, _
        includeForm As Boolean) As Collection

    Dim Module      As VBComponent
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim ProcedureName As String
    Dim i           As Long
    For Each Module In oProject.VBComponents
        If Not includeModule And Module.Type = vbext_ct_StdModule Then GoTo SKIP
        If Not includeClass And Module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If Not includeDocument And Module.Type = vbext_ct_Document Then GoTo SKIP
        If Not includeForm And Module.Type = vbext_ct_MSForm Then GoTo SKIP
        With Module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .countOfLines

                For i = 0 To 4
                    procKind = i
                    ProcedureName = .ProcOfLine(lineNum, procKind)
                    If ProcedureName <> vbNullString Then

                        'Possible Error:
                        '---------------
                        ' Procedures with an underscore _ which is used in events
                        ' and Private Procedures
                        ' may have the same name in different components

                        If InStr(1, ProcedureName, "_") = 0 Then
                            ProcedureName = .ProcOfLine(lineNum, procKind)
                            On Error Resume Next
                            coll.Add ProcedureName, Module.Name & "|" & ProcedureName & "|" & procKind
                            On Error GoTo 0
                        End If

                        Exit For

                    End If
                Next
                lineNum = .procStartLine(ProcedureName, procKind) + .ProcCountLines(ProcedureName, procKind) + 1
            Loop
        End With
SKIP:
    Next Module

    Set ProceduresNames = coll
End Property

Function ProceduresArray()
    Dim Module      As VBComponent
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim ProcedureName As String
    Dim cProcedure  As aProcedure
    Dim arr
    Dim i           As Long
    Dim BodyLine    As Long
    ReDim arr(1 To 4, 1 To 1)
    For Each Module In oProject.VBComponents
        With Module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .countOfLines
                If arr(UBound(arr, 1), 1) <> "" Then
                    ReDim Preserve arr(1 To 4, 1 To UBound(arr, 2) + 1)
                End If
                On Error Resume Next
                For i = 0 To 4
                    procKind = i
                    ProcedureName = .ProcOfLine(lineNum, procKind)
                    If ProcedureName <> vbNullString Then
                        BodyLine = Module.CodeModule.ProcBodyLine(ProcedureName, procKind)
                        arr(1, UBound(arr, 2)) = Switch(Module.Type = vbext_ct_StdModule, "Module", _
                                Module.Type = vbext_ct_ClassModule, "Class", _
                                Module.Type = 3, "Userform", _
                                Module.Type = 100, "Document", _
                                Module.Type = 11, "ActiveXDesigner")
                        arr(2, UBound(arr, 2)) = Module.Name
                        arr(3, UBound(arr, 2)) = Switch(procKind = vbext_pk_Proc, IIf(Module.CodeModule.lines(BodyLine, 1) Like "*Sub " & ProcedureName & "*", "Sub", "Function"), _
                                procKind = vbext_pk_Let, "Let", _
                                procKind = vbext_pk_Set, "Set", _
                                procKind = vbext_pk_Get, "Get")
                        arr(4, UBound(arr, 2)) = ProcedureName
                        Exit For
                    End If
                Next
                On Error GoTo 0
                lineNum = .procStartLine(ProcedureName, procKind) + .ProcCountLines(ProcedureName, procKind) + 1
            Loop
        End With
SKIP:
    Next Module
    ProceduresArray = WorksheetFunction.Transpose(arr)
End Function

Public Sub ExportProcedures(targetPath As String, createModuleSubfolders As Boolean)
    targetPath = VBA.Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim ams As aModules: Set ams = aModules.Initialize(oProject)
    ams.ExportProcedures targetPath, createModuleSubfolders
End Sub

Public Sub ExportModules(targetPath As String)
    targetPath = VBA.Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim ams As aModules: Set ams = aModules.Initialize(oProject)
    ams.Export targetPath, createSubfolders:=True
End Sub

Public Sub ExportXML(targetPath As String)
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim TmpFile     As String
    TmpFile = WorkbookObject.path & "\temp_workbook_file" & Extension
    WorkbookObject.SaveCopyAs TmpFile
    Dim c           As New clsEditOpenXML
    c.ExtractRibbonX TmpFile, targetPath & "customUI.xml"
    Kill TmpFile
    FolderDelete WorkbookObject.path & "\Unzipped " & "temp_workbook_file" & Extension & ".zip"
    Set c = Nothing
End Sub

Public Sub ExportCodeUnified(targetPath As String)
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim result As String, tmp As String
    Dim am          As aModule
    For Each am In aModules.Initialize(oProject).Items
        tmp = "'" & am.Name & vbTab & am.TypeToString & vbNewLine & vbNewLine & am.code
        result = IIf(result = "", tmp, result & vbNewLine & vbNewLine & tmp)
    Next
    TxtOverwrite targetPath & "#UnifiedProject.txt", result
End Sub







Sub ModulesMerge(TargetModule As VBComponent, ParamArray Modules() As Variant)
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE CLASS aModule
    Dim element     As Variant
    Dim Module      As VBComponent
    For Each element In Modules
        Set Module = element
        If Module.Type = vbext_ct_StdModule Then
            If Module.Name <> TargetModule.Name Then
                aModule.Initialize(Module).CodeMove TargetModule
            End If
        End If
    Next
End Sub




Sub CreateLinkedTableSheet()
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = CreateOrSetSheet("Linked_Table_" & NameClean, WorkbookObject)
    If TargetWorksheet.Range("A1") <> "" Then Exit Sub
    Dim var
    ReDim var(1 To 8)

    var(1) = "WORKBOOK"
    var(2) = "MODULE TYPE"
    var(3) = "MODULE NAME"
    var(4) = "PROCEDURE"
    var(5) = "LINKED PROCEDURES"
    var(6) = "LINKED CLASSES"
    var(7) = "LINKED USERFORMS"
    var(8) = "LINKED DECLARATIONS"

    With TargetWorksheet
        .Range("A1:H1").value = var
        With .rows(1).Cells.Font
            .Bold = True
            .Size = 14
        End With
    End With

End Sub

Sub CreateLinkedTable()

    StartTimer "tableOfProcedures"

    Declarations.createTable

    CreateLinkedTableSheet

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = WorkbookObject.Sheets("Linked_Table_" & NameClean)
    On Error Resume Next
    TargetWorksheet.ListObjects(1).Unlist
    On Error GoTo 0
    TargetWorksheet.Range("A1").CurrentRegion.offset(1).clear

    Dim var
    ReDim var(1 To 8, 1 To 1)

    Dim counter     As Long
    counter = 1

    Dim Module      As VBComponent
    Dim Procedures  As Collection

    Dim Procedure   As aProcedure
    For Each Module In oProject.VBComponents
        Select Case Module.Type
            Case vbext_ct_StdModule, vbext_ct_Document, vbext_ct_MSForm    '@TODO check if adding these is ok
                Set Procedures = aModule.Initialize(Module).Procedures
                For Each Procedure In Procedures
                    ReDim Preserve var(1 To 8, 1 To counter)
                    var(1, counter) = WorkbookObject.Name
                    var(2, counter) = aModule.Initialize(Module).TypeToString
                    var(3, counter) = Module.Name
                    var(4, counter) = Procedure.Name
                    var(5, counter) = aCollection.Init(Procedure.Dependencies.LinkedProcedures).ToString(vbNewLine)
                    var(6, counter) = aCollection.Init(Procedure.Dependencies.LinkedClasses).ToString(vbNewLine)
                    var(7, counter) = aCollection.Init(Procedure.Dependencies.LinkedUserforms).ToString(vbNewLine)
                    var(8, counter) = aCollection.Init(Procedure.Dependencies.LinkedDeclarations).ToString(vbNewLine)

                    counter = counter + 1
                Next
        End Select
    Next

    var = WorksheetFunction.Transpose(var)
    ArrayToRange2D var, TargetWorksheet.Range("A2")

    LinkedProceduresTableSort
    TargetWorksheet.ListObjects.Add xlSrcRange, TargetWorksheet.Range("A1").CurrentRegion, , xlYes
    TargetWorksheet.Cells.VerticalAlignment = xlVAlignTop

    EndTimer

End Sub

Private Sub LinkedProceduresTableSort()
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = WorkbookObject.Worksheets("Linked_Table_" & NameClean)
    Dim sort1       As String
    sort1 = "A1"
    Dim sort2       As String
    sort2 = "C1"
    Dim sort3       As String
    sort3 = "D1"

    With TargetWorksheet.Sort
        .SortFields.clear
        .SortFields.Add Key:=TargetWorksheet.Range(sort1), Order:=xlAscending
        If Not sort2 = "" Then
            .SortFields.Add Key:=TargetWorksheet.Range(sort2), Order:=xlAscending
        End If
        If Not sort3 = "" Then
            .SortFields.Add Key:=TargetWorksheet.Range(sort3), Order:=xlAscending
        End If

        .SetRange TargetWorksheet.Range("A1").CurrentRegion
        '        .header = xlYes
        .Apply
    End With
End Sub
