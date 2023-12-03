VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Body            As String               :
' Fun ClassCalls      As Variant              :
' Fun Code            As String               :
' Fun Component       As VBComponent          :
' Fun Contains        As Variant              :
' Fun Copy            As Boolean              :
' Fun Duplicate       As Boolean              :
' Fun Enums           As aModuleEnums         :
' Fun Extension       As String               :
' Fun Folders         As aModuleFolders       :
' Fun Header          As String               :
' Fun HeaderContains  As Boolean              :
' Fun Ignore          As Boolean              :
' Fun Initialize      As aModule              :
' Fun LineLike        As Long                 :
' Fun LinesLike       As Collection           :
' Fun ListOfInclude   As Collection           :
' Fun Name            As String               :
' Fun Procedures      As aModuleProcedures    :
' Fun TodoList        As Variant              :
' Fun TypeToLong      As vbext_ComponentType  :
' Fun TypeToString    As String               :
' Fun Types           As aModuleTypes         :
' Get Active          As aModule              :
' Get Project         As VBProject            :
' Get WorkbookObject  As Workbook             :
' Sub Activate                                :
' Sub CodeMove                                :
' Sub CodeRemove                              :
' Sub CommentsRemove                          :
' Sub CommentsToOwnLine                       :
' Sub Delete                                  :
' Sub DeleteIfEmpty                           :
' Sub DisableDebugPrint                       :
' Sub DisableStop                             :
' Sub EnableDebugPrint                        :
' Sub EnableStop                              :
' Sub Export                                  :
' Sub HeaderAdd                               :
' Sub Indent                                  :
' Sub PredeclaredId                           :
' Sub PrintListOfInclude                      :
' Sub PrintTodoList                           :
' Sub ProcedureFoldDeclarations               :
' Sub RemoveEmptyLines                        :
' Sub RemoveEmptyLinesButLeaveOne             :
' Sub RemoveLinesLike                         :
' Sub Rename                                  :
'--------------------------------------------------
'@EndFolder ListOfProcedures

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aModule
'* Purpose    :
'* Copyright  :
'*
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 29-06-2023 14:39    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Private oModule As VBComponent

Public Function Enums() As aModuleEnums
    Set Enums = New aModuleEnums
    Enums.Initialize Me
End Function

Public Function Types() As aModuleTypes
    Set Types = New aModuleTypes
    Types.Initialize Me
End Function

Public Function Procedures() As aModuleProcedures
    Set Procedures = New aModuleProcedures
    Procedures.Initialize Me
End Function
Public Function Folders() As aModuleFolders
    Set Folders = New aModuleFolders
    Folders.Initialize Me
End Function

Public Function Initialize(Module As VBComponent) As aModule
    Set oModule = Module
    Set Initialize = Me
End Function

Public Property Get Active() As aModule
    If Not Application.VBE.ActiveCodePane Is Nothing Then
        Set Active = Initialize(Application.VBE.ActiveCodePane.CodeModule.Parent)
    End If
End Property

Public Function Component() As VBComponent
    Set Component = oModule
End Function

Public Property Get WorkbookObject() As Workbook
    Dim s           As String
    s = oModule.Collection.Parent.fileName
    s = Right(s, Len(s) - InStrRev(s, "\"))
    Set WorkbookObject = Workbooks(s)
End Property
Public Property Get Project() As VBProject
    Set Project = oModule.Collection.Parent
End Property
Function Name() As String
    Name = oModule.Name
    If oModule.Type = vbext_ct_Document And oModule.Name <> "ThisWorkbook" Then
        On Error GoTo HELL
        Name = GetSheetByCodeName(WorkbookObject, oModule.Name).Name
        On Error GoTo 0
    End If
    Exit Function
HELL:
    Debug.Print "Probably the file is in another language"
    Name = "ThisWorkbook"
End Function


Sub Activate()
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    With Application.VBE.MainWindow
        .Visible = True
        .WindowState = vbext_ws_Maximize
    End With
    With oModule.CodeModule.CodePane
        .Show
        .Window.Visible = True
        .Window.WindowState = vbext_ws_Maximize
        .Window.SetFocus
        .SetSelection 1, 1, 1, 1
    End With
End Sub

Sub Export(targetPath As String)
    If targetPath = "" Then Exit Sub    '<- FOR when called with FolderPicker
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    oModule.Export targetPath & Name & Extension
End Sub

Function code() As String
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    With oModule.CodeModule
        If .countOfLines = 0 Then code = "": Exit Function
        code = .lines(1, .countOfLines)
    End With
End Function

Public Function Body() As String
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    With oModule.CodeModule
        If .countOfLines - .CountOfDeclarationLines = 0 Then Exit Function
        Body = .lines(.CountOfDeclarationLines + 1, .countOfLines - .CountOfDeclarationLines)
    End With
End Function

Function Contains( _
        this As String, _
        WholeWord As Boolean, _
        MatchCase As Boolean, _
        PatternSearch As Boolean)
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule

    If oModule.CodeModule.countOfLines = 0 Then Exit Function
    Contains = oModule.CodeModule.Find( _
            this, _
            1, _
            1, _
            oModule.CodeModule.countOfLines, _
            10000, _
            WholeWord, _
            MatchCase, _
            PatternSearch)
End Function

Function Header() As String
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.CountOfDeclarationLines = 0 Then Exit Function
    With oModule.CodeModule
        Header = .lines(1, .CountOfDeclarationLines)
    End With
End Function

Function HeaderContains(this As String) As Boolean
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    HeaderContains = (InStr(1, Header, this) > 0)
End Function

Function Ignore() As Boolean
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    Ignore = InStr(1, Header, "@IGNORE", vbTextCompare) > 0
End Function

Sub CodeRemove()
    If Ignore Then Exit Sub
    If oModule.CodeModule.countOfLines = 0 Then Exit Sub
    oModule.CodeModule.DeleteLines 1, oModule.CodeModule.countOfLines    '+ 1
End Sub

Sub CodeMove(ToModule As VBComponent)
    If Ignore Then Exit Sub
    With ToModule.CodeModule
        .InsertLines 1, Header
        .InsertLines .countOfLines + 1, Body
    End With
    CodeRemove
End Sub

Function TypeToLong() As vbext_ComponentType
    TypeToLong = oModule.Type
End Function

Function TypeToString() As String
    Select Case oModule.Type
        Case vbext_ct_ActiveXDesigner
            TypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            TypeToString = "Class"
        Case vbext_ct_Document
            TypeToString = "Document"
        Case vbext_ct_MSForm
            TypeToString = "UserForm"
        Case vbext_ct_StdModule
            TypeToString = "Module"
        Case Else
            TypeToString = "Unknown Type: " & CStr(oModule.Type)
    End Select
End Function

Public Sub PredeclaredId(Optional Toggle As Boolean = True)
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE PROCEDURE TxtRead
    '@INCLUDE CLASS aModule
    
    If Not oModule.Type = vbext_ct_ClassModule Then Exit Sub
    
    Dim FName1      As String: FName1 = ThisWorkbook.path & "\" & "TempClass.cls"
    oModule.Export FName1
    TxtOverwrite FName1, VBA.Replace( _
            VBA.Replace(TxtRead(FName1), _
            IIf(Toggle = True, _
            "Attribute VB_PredeclaredId = False", _
            "Attribute VB_PredeclaredId = True"), _
            IIf(Toggle = True, _
            "Attribute VB_PredeclaredId = True", _
            "Attribute VB_PredeclaredId = False"), _
            1, _
            1), _
            vbLf, vbCrLf, _
            1, _
            -1)

    Delete
    Project.VBComponents.Import FName1
    Kill FName1
    Activate
End Sub

Sub Delete()
    If Ignore Then Exit Sub
    If oModule.Type = vbext_ct_Document Then
        Exit Sub
    Else
        Application.DisplayAlerts = False
        Dim targetFolder As String: targetFolder = ThisWorkbook.path & "\backup\"
        FoldersCreate targetFolder
        oModule.Export targetFolder & oModule.Name: Debug.Print "kept backup in " & targetFolder
        WorkbookObject.VBProject.VBComponents.Remove oModule
        Application.DisplayAlerts = True
    End If
End Sub

'@TODO fix this
Public Function Duplicate( _
        Optional NewName As String, _
        Optional ActivateNewModule As Boolean = True) As Boolean
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    Dim oldName As String
    oldName = oldName = Name
    If NewName = "" Then NewName = Name & "_Copy"
    On Error GoTo Error_Handler
    Dim targetFile  As String
    targetFile = Environ("Temp") & "\" & NewName & Extension
    oModule.Name = NewName
    oModule.Export targetFile
    oModule.Name = oldName
    Project.VBComponents.Import targetFile
    If Len(Dir(targetFile)) > 0 Then Kill Replace(targetFile, Extension, ".*")
    If ActivateNewModule = True Then Project.VBComponents(NewName).Activate
    '    DuplicateUserForm = True
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: aModule.Duplicate" & vbCrLf & _
            "Error Description: " & Err.Description & _
            Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
            , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function


Public Sub Rename(NewName As String)
    If Not ModuleExists(NewName, WorkbookObject) Then oModule.Name = NewName
End Sub

Function Extension() As String
    Dim ext         As String
    Select Case oModule.Type
        Case vbext_ct_MSForm
            ext = ".frm"
        Case vbext_ct_ClassModule, vbext_ct_Document
            ext = ".cls"
            '        Case vbext_ct_Document
            '            If oModule.Name = "ThisWorkbook" Then
            '                Extension = ".WBcls"
            '            Else
            '                Extension = ".DOCcls"
            '            End If
        Case vbext_ct_StdModule
            ext = ".bas"
        Case Else
            ext = ".txt"
    End Select
    Extension = ext
End Function

Public Function Copy( _
        TargetWorkbook As Workbook, _
        OverwriteExisting As Boolean) As Boolean

    If Name = "ThisWorkbook" Then Exit Function
    If oModule.Type = vbext_ct_Document Then Exit Function
    If WorkbookObject.Name = TargetWorkbook.Name Then Exit Function
    Dim TempModule  As VBIDE.VBComponent

    If ModuleExists(Name, TargetWorkbook) Then
        If OverwriteExisting = True Then
            With TargetWorkbook.VBProject
                .VBComponents.Remove .VBComponents(Name)
            End With
        Else
            Exit Function
        End If
    End If

    Dim FName       As String
    FName = Environ("Temp") & "\" & Name & Extension
    oModule.Export fileName:=FName

    TargetWorkbook.VBProject.VBComponents.Import fileName:=FName
    Kill FName
    Copy = True
End Function








Public Function ClassCalls() As Variant
    'classCallsOfModule(0) is the class name
    'classCallsOfModule(1) is the keyword for the class name (eg dim clsCal as new classCalendar)
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ClassNames
    '@INCLUDE PROCEDURE classCallsOfModule
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules

    Dim code        As Variant
    Dim element     As Variant
    Dim Keyword     As Variant
    Dim var         As Variant
    ReDim var(1 To 2, 1 To 1)
    Dim counter     As Long: counter = 0
    Dim ClassName   As Variant
    If oModule.CodeModule.CountOfDeclarationLines > 0 Then
        code = oModule.CodeModule.lines(1, oModule.CodeModule.CountOfDeclarationLines)
        code = Replace(code, "_" & vbNewLine, "")
        code = Split(code, vbNewLine)
        code = Filter(code, " As ", , vbTextCompare)
        For Each element In code
            element = Trim(element)
            If element Like "* As *" Then
                Keyword = Split(element, " As ")(0)
                Keyword = Split(Keyword, " ")(UBound(Split(Keyword, " ")))
                element = Split(element, " As ")(1)
                element = Replace(element, "New ", "")
                For Each ClassName In aModules.Initialize(Project).ClassNames
                    If element = ClassName Then
                        ReDim Preserve var(1 To 2, 1 To counter + 1)
                        var(1, UBound(var, 2)) = element
                        var(2, UBound(var, 2)) = Keyword
                        counter = counter + 1
                        '                        Exit For
                    End If
                Next
            End If
        Next
        If var(1, 1) <> "" Then
            If UBound(var, 2) > 1 Then
                ClassCalls = WorksheetFunction.Transpose(var)
            Else
                Dim VAR2(1 To 1, 1 To 2)
                VAR2(1, 1) = var(1, 1)
                VAR2(1, 2) = var(2, 1)
                ClassCalls = VAR2
            End If
        End If
    End If

End Function


Sub PrintListOfInclude()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE dp
    '@INCLUDE CLASS aModule
    dp ListOfInclude
End Sub

Function ListOfInclude() As Collection
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ArrayQuickSort
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aCollection

    Dim var: var = Split(code, vbNewLine)
    ArrayQuickSort Filter(Filter(var, "'@INCLUDE ", True), """", False)
    Dim coll        As New Collection
    Dim el
    On Error Resume Next
    For Each el In var
        el = Trim(el)
        If el Like "'@INCLUDE *" Then
            coll.Add el, CStr(el)
        End If
    Next
    On Error GoTo 0
    var = aCollection.Init(coll).ToArray
    Dim matches     As String: matches = Join(var, vbNewLine)
    Dim out         As String: out = out & vbNewLine & "'=== Module: " & oModule.Name & "===" & vbNewLine & matches
    Set ListOfInclude = coll
End Function


Sub PrintTodoList()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE dp
    '@INCLUDE CLASS aModule
    dp TodoList
End Sub



















'* @TODO Created: 20-11-2023 14:24 Author: Anastasiou Alex
'* @TODO add this as aProcedure property so as to list Module > Procedure > Line - description

Function TodoList()
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    Dim out         As String
    Dim matches     As String
    matches = Join(Filter(Split(code, vbNewLine), "@TODO", True), vbNewLine)
    out = out & vbNewLine & "'---------"
    out = out & vbNewLine & "'Module: " & Name
    out = out & vbNewLine & "'---------"
    '    out = out & vbNewLine & "'" & Replace(Matches, vbNewLine, vbNewLine & "'")
    out = out & vbNewLine & matches    ' "'" & Replace(Matches, vbNewLine, vbNewLine & "'")
    TodoList = out
End Function


Sub HeaderAdd()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ActiveProcedure
    '@INCLUDE PROCEDURE PadRight
    '@INCLUDE CLASS aModule
    '@INCLUDE DECLARATION AUTHOR_COPYRIGHT
    '@INCLUDE DECLARATION AUTHOR_EMAIL
    '@INCLUDE DECLARATION AUTHOR_NAME
    '@INCLUDE DECLARATION AUTHOR_OTHERTEXT

    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab & vbTab & vbTab & vbTab
    Const sFersLine As String = "'* Modified   :" & vbTab & "Date and Time" & vbTab & vbTab & "Author" & vbTab4 & "Description" & vbCrLf

    Dim txtName     As String: txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")
    txtName = "'* Author     :" & vbTab & txtName & vbCrLf

    Dim txtContacts As String: txtContacts = AUTHOR_EMAIL
    If txtContacts <> vbNullString Then txtContacts = "'* Contacts   :" & vbTab & txtContacts & vbCrLf

    Dim txtCopyright As String: txtCopyright = AUTHOR_COPYRIGHT
    '    If txtCopyright <> vbNullString Then txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf
    txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf

    Dim txtOther    As String: txtOther = AUTHOR_OTHERTEXT
    If txtOther <> vbNullString Then txtOther = "'* Note       :" & vbTab & txtOther & vbCrLf

    Dim txtMedia    As String: txtMedia = "'* " & vbLf & AUTHOR_MEDIA
    Dim CurentCodePane As CodePane: Set CurentCodePane = oModule.CodeModule.CodePane
    Dim i           As Byte
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim sProc       As String: sProc = ActiveProcedure
    Dim nLine       As Long: nLine = 1
    Dim sTemp       As String
    Dim sTime       As String
    Dim sType       As String
    Dim sProcDeclartion As String
    Dim sProcArguments As String

    Dim sUser       As String
    On Error GoTo ErrorHandler

    sUser = "Alex"
    If sUser = vbNullString Then sUser = Environ("UserName")

    sTemp = Replace(String(90, "*"), "**", "* ")
    sTime = Format(Now, "dd-mm-yyyy hh:nn")
    sProc = oModule.Name
    sType = "'* " & PadRight(Me.TypeToString, 11) & ":" & vbTab & sProc & vbCrLf

    If LineLike("* " & PadRight(Me.TypeToString, 11) & ": " & oModule.Name, True, True) > 0 Then GoTo UpdateHeader

    With oModule.CodeModule.CodePane
        sTemp = vbLf & "'" & sTemp & vbCrLf & _
                sType & _
                "'* Purpose    :" & vbCrLf & _
                txtCopyright & "'* " & vbCrLf & _
                txtName & _
                txtContacts & _
                txtOther & _
                txtMedia & _
                sFersLine & _
                "'* Created    :" & vbTab & sTime & vbTab & sUser & vbCrLf & _
                sProcArguments & _
                "'" & sTemp
        sTemp = sTemp & vbNewLine & IIf(HeaderContains("Option Explicit"), "", "Option Explicit" & vbNewLine) & _
                                    IIf(HeaderContains("Option Private Module"), "", "Option Private Module" & vbNewLine) & _
                                    IIf(HeaderContains("Option Base 1"), "", "Option Base 1") & vbNewLine & _
                                    IIf(HeaderContains("C_MODULE_NAME"), "", "C_MODULE_NAME = " & """" & Me.Name & """" & vbNewLine)

'         sTemp = sTemp & "'You may later want to use: amodule.active.ListProceduresPublic"
        .CodeModule.InsertLines 1, sTemp & vbNewLine
    End With

NormalExit:
    Exit Sub

UpdateHeader:
    nLine = LineLike("'* Created    : *", True, True) + 1
    'put update to bottom of changes list
    Do While Trim(oModule.CodeModule.lines(nLine, 1)) Like "'* Modified*)"
        nLine = nLine + 1
    Loop
    oModule.CodeModule.InsertLines nLine, "'* Modified   :" & vbTab & sTime & vbTab & sUser
    GoTo NormalExit

ErrorHandler:
    GoTo NormalExit
End Sub




Public Function LineLike( _
        this As String, _
        IgnoreIndentation As Boolean, _
        IgnoreCase As Boolean) As Long
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.countOfLines = 0 Then Exit Function
    If IgnoreCase Then this = UCase(this)
    Dim sLine       As String
    Dim i           As Long
    For i = 1 To oModule.CodeModule.countOfLines
        sLine = oModule.CodeModule.lines(i, 1)
        If IgnoreIndentation Then sLine = Trim(sLine)
        If IgnoreCase Then sLine = UCase(sLine)
        If sLine Like this Then
            LineLike = i
            Exit Function
        End If
    Next
End Function

Public Function LinesLike( _
        this As String, _
        IgnoreIndentation As Boolean, _
        IgnoreCase As Boolean) As Collection
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.countOfLines = 0 Then Exit Function
    If IgnoreCase Then this = UCase(this)
    Dim coll        As New Collection
    Dim sLine       As String
    Dim i           As Long
    For i = 1 To oModule.CodeModule.countOfLines
        sLine = oModule.CodeModule.lines(i, 1)
        If IgnoreIndentation Then sLine = Trim(sLine)
        If IgnoreCase Then sLine = UCase(sLine)
        If sLine Like this Then
            coll.Add i
        End If
    Next
    Set LinesLike = coll

End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 13:56    Alex                (aModule.cls > RemoveLinesLike)

Public Sub RemoveLinesLike(this As String)
'@LastModified 2308221356
    Dim i As Long
    For i = oModule.CodeModule.countOfLines To 1 Step -1
        If oModule.CodeModule.lines(i, 1) Like this Then
            oModule.CodeModule.DeleteLines i, 1
        End If
    Next
End Sub


Sub EnableDebugPrint()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Format.EnableDebugPrint
    Next
End Sub
Sub DisableDebugPrint()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Format.DisableDebugPrint
    Next
End Sub
Public Sub EnableStop()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Format.EnableStop
    Next
End Sub
Public Sub DisableStop()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Format.DisableStop
    Next
End Sub

Public Sub ProcedureFoldDeclarations()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Format.FoldDeclaration
    Next
End Sub


Sub CommentsToOwnLine()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE CommentsMoveToOwnLine
    '@INCLUDE CLASS aModule
    Dim s           As String
    With oModule.CodeModule
        If .countOfLines = 0 Then Exit Sub
        s = CommentsMoveToOwnLine(code)
        CodeRemove
        .AddFromString s
    End With
End Sub

Public Sub Indent()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE IsBlockEnd
    '@INCLUDE PROCEDURE IsBlockStart
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.countOfLines = 0 Then Exit Sub
    Dim LineStart As Long, LineEnd As Long
    LineStart = 1
    LineEnd = oModule.CodeModule.countOfLines
    Dim nIndent     As Long
    Dim s           As String
    Dim i           As Long
    For i = LineStart To LineEnd
        s = LTrim$(oModule.CodeModule.lines(i, 1))
        If Trim(s) <> "" Then
            If IsBlockEnd(s) Then nIndent = nIndent - 1
            If nIndent < 0 Then nIndent = 0
            s = Space$(nIndent * 4) & s
            oModule.CodeModule.ReplaceLine i, s
            If IsBlockStart(LTrim$(s)) Then nIndent = nIndent + 1
        End If
    Next
End Sub

Public Sub RemoveEmptyLines()
    Dim code        As String
    Dim currentline As String, nextline As String
    Dim i           As Long
    With oModule.CodeModule
        If .countOfLines > 0 Then
            For i = .countOfLines - 1 To 1 Step -1
                currentline = Trim(.lines(i, 1))
                If Len(currentline) = 0 Then
                    .DeleteLines i, 1
                End If
            Next
        End If
    End With
End Sub

Public Sub CommentsRemove()
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Format.CommentsRemove ap.lines.FirstOfBody, , False
    Next ap
End Sub










Sub DeleteIfEmpty()
    If Ignore Then Exit Sub
    If Name = "ThisWorkbook" Then Exit Sub
    If oModule.CodeModule.countOfLines = 0 Then Delete
End Sub

Public Sub RemoveEmptyLinesButLeaveOne()
    Dim code        As String
    Dim currentline As String, nextline As String
    Dim i           As Long
    With oModule.CodeModule
        If .countOfLines > 0 Then
            For i = .countOfLines - 1 To 1 Step -1
                currentline = Trim(.lines(i, 1))
                nextline = Trim(.lines(i + 1, 1))
                If Len(currentline) = 0 Then
                    If Len(nextline) = 0 Then
                        .DeleteLines i + 1, 1
                    End If
                End If
            Next
        End If
    End With
End Sub
