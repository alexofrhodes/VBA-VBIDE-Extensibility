VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureDependencies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun CallerModules              As Collection              :
' Fun CallerModulesToString      As String                  :
' Fun Callers                    As Collection              :
' Fun CallersToString            As String                  :
' Fun DeclarationsTableKeywords  As Collection              :
' Fun Initialize                 As aProcedureDependencies  :
' Fun LinkedClasses              As Collection              :
' Fun LinkedDeclarations         As Collection              :
' Fun LinkedProcedures           As Collection              :
' Fun LinkedProceduresDeep       As Collection              :
' Fun LinkedSheets               As Collection              :
' Fun LinkedUserforms            As Variant                 :
' Fun collLinkedProcedures       As Collection              :
' Fun collLinkedProceduresDeep   As Collection              :
' Sub AddToLinkedTable                                      :
' Sub BringLinkedProceduresHere                             :
' Sub BringProcedureHere                                    :
' Sub Export                                                :
' Sub ExportDeclaration                                     :
' Sub ExportLinkedCode                                      :
' Sub ImportClass                                           :
' Sub ImportDeclaration                                     :
' Sub ImportDependencies                                    :
' Sub ImportProcedure                                       :
' Sub ImportUserform                                        :
' Sub InjectLinkedClasses                                   :
' Sub InjectLinkedDeclarations                              :
' Sub InjectLinkedLists                                     :
' Sub InjectLinkedProcedures                                :
' Sub InjectLinkedUserforms                                 :
' Sub RemoveIncludeLines                                    :
' Sub Update                                                :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Private coll_LinkedProcedures As Collection
Private coll_LinkedProceduresDeep As Collection

Public Function Initialize(Procedure As aProcedure) As aProcedureDependencies
    Set oProcedure = Procedure
    Set coll_LinkedProcedures = Nothing
    Set coll_LinkedProceduresDeep = Nothing
    Set Initialize = Me
End Function

Public Function collLinkedProceduresDeep() As Collection
    If coll_LinkedProceduresDeep Is Nothing Then
        Set coll_LinkedProceduresDeep = LinkedProceduresDeep
    End If
    Set collLinkedProceduresDeep = coll_LinkedProceduresDeep
End Function

Public Function collLinkedProcedures() As Collection
    If coll_LinkedProcedures Is Nothing Then
        Set coll_LinkedProcedures = LinkedProcedures
    End If
    Set collLinkedProcedures = coll_LinkedProcedures
End Function

Public Function LinkedProcedures() As Collection
    Dim Procedures  As Collection
    Set Procedures = aProject.Initialize(oProcedure.oWorkbook).ProceduresNames(True, False, False, False)
    Dim S           As String: S = oProcedure.Code.All
    Dim coll        As New Collection
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) <> UCase(CStr(oProcedure)) Then
            If RegexTest(S, "\W" & CStr(Procedure) & "[.(\W]") = True Then
                coll.Add Procedure, CStr(Procedure)
            End If
        End If
    Next
    Set LinkedProcedures = coll
End Function

Public Function LinkedProceduresDeep() As Collection
    Dim AllProcedures As Collection: Set AllProcedures = aProject.Initialize(oProcedure.oWorkbook).ProceduresNames(True, False, False, False)
    Dim Processed   As Collection: Set Processed = New Collection
    Dim CalledProcedures As Collection: Set CalledProcedures = New Collection

    Dim Procedure   As Variant
    Dim Module      As VBComponent

    Processed.Add CStr(oProcedure), CStr(oProcedure)
    CalledProcedures.Add CStr(oProcedure), CStr(oProcedure)
    On Error Resume Next
    For Each Procedure In LinkedProcedures
        CalledProcedures.Add CStr(Procedure), CStr(Procedure)
    Next
    On Error GoTo 0

    Dim ap          As aProcedure
    Dim CalledProceduresCount As Long
    CalledProceduresCount = CalledProcedures.count
    Dim element
repeat:
    For Each element In CalledProcedures
        If Not aCollection.Init(Processed).Contains(, CStr(element)) Then
            On Error Resume Next
            Set ap = New aProcedure
            For Each Procedure In ap.Initialize(oProcedure.oWorkbook, , CStr(element)).Dependencies.LinkedProcedures
                CalledProcedures.Add CStr(Procedure), CStr(Procedure)
            Next
            On Error GoTo 0
            Processed.Add CStr(element), CStr(element)
        End If
    Next
    If CalledProcedures.count > CalledProceduresCount Then
        CalledProceduresCount = CalledProcedures.count
        GoTo repeat
    End If

    Set LinkedProceduresDeep = aCollection.Init(CalledProcedures).Sort.Items
End Function

Public Sub Export()
    Dim Procedure
    For Each Procedure In collLinkedProceduresDeep
        LinkedExport CStr(Procedure)
    Next
End Sub

Private Sub LinkedExport(Procedure As String)
    Dim Module      As VBComponent
    Set Module = ModuleOfProcedure(oProcedure.oWorkbook, Procedure)
    Dim c           As New aProcedure
    c.Initialize oProcedure.oWorkbook, Module, Procedure
    Dim proclastmod As String: proclastmod = c.CustomProperties.LastModified
    If proclastmod = vbNullString Then
        c.Dependencies.InjectLinkedLists
        c.CustomProperties.LastModified = Format(Now, "yymmddhhnn")
        proclastmod = c.CustomProperties.LastModified
    End If
    Dim S           As String: S = c.Code.All
    Dim FileFullName As String: FileFullName = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"

    If FileExists(FileFullName) Then
        Dim filelastmod
        filelastmod = StringLastModified(TxtRead(FileFullName))
        If proclastmod > filelastmod Then
            Debug.Print "OVERWROTE " & Procedure
            TxtOverwrite FileFullName, S
        Else
            Debug.Print "Skipping " & Procedure & " because:"
            Debug.Print vbTab & "StringLastModified(TxtRead(FileFullName)) >= Property(Modified)"
        End If
    Else
        Debug.Print "NEW " & Procedure
        TxtOverwrite FileFullName, S
    End If
    Dim element
    For Each element In c.Dependencies.LinkedUserforms
        oProcedure.oWorkbook.VBProject.VBComponents(element).Export LOCAL_LIBRARY_USERFORMS & element & ".frm"
    Next
    For Each element In c.Dependencies.LinkedClasses
        oProcedure.oWorkbook.VBProject.VBComponents(element).Export LOCAL_LIBRARY_CLASSES & element & ".cls"
    Next
    For Each element In c.Dependencies.LinkedDeclarations
        oProcedure.Dependencies.ExportDeclaration CStr(element)
    Next
End Sub

Sub ExportDeclaration(DeclarationName As String)
    aProject.Initialize(oProcedure.oWorkbook).Declarations.createTable
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = oProcedure.oWorkbook.Sheets("Declarations_Table")

    Dim codeName    As String
    Dim codeText    As String
    Dim cell        As Range
    On Error Resume Next
    Set cell = TargetWorksheet.Columns(3).Find(DeclarationName, LookAt:=xlWhole)
    On Error GoTo 0
    If cell Is Nothing Then Exit Sub

    codeName = DeclarationName
    codeText = cell.offset(0, 1).text
    TxtOverwrite LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt", codeText

End Sub
Public Sub InjectLinkedLists()
    RemoveIncludeLines
    InjectLinkedDeclarations
    InjectLinkedClasses
    InjectLinkedUserforms
    InjectLinkedProcedures
End Sub


Public Sub RemoveIncludeLines()
    oProcedure.Format.RemoveLinesLike "'@INCLUDE *"
End Sub

Public Sub InjectLinkedClasses()
    Dim ListOfImports As String
    Dim S           As String: S = oProcedure.Code.All
    Dim element     As Variant
    For Each element In LinkedClasses
        If InStr(1, S, "@INCLUDE CLASS " & element) = 0 _
                And InStr(1, ListOfImports, "@INCLUDE CLASS " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE CLASS " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE CLASS " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oProcedure.oModule.CodeModule.InsertLines _
                oProcedure.lines.FirstOfBodyAfterComments, ListOfImports
    End If
End Sub

Public Sub InjectLinkedProcedures()
    Dim Procedures  As Collection
    Set Procedures = collLinkedProcedures
    Dim ListOfImports As String
    Dim S           As String: S = oProcedure.Code.All
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If InStr(1, S, "@INCLUDE PROCEDURE " & Procedure) = 0 And InStr(1, ListOfImports, "@INCLUDE PROCEDURE " & Procedure) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE PROCEDURE " & Procedure
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE PROCEDURE " & Procedure
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oProcedure.oModule.CodeModule.InsertLines oProcedure.lines.FirstOfBodyAfterComments, ListOfImports
    End If
End Sub

Public Sub InjectLinkedUserforms()
    Dim ListOfImports As String
    Dim S           As String: S = oProcedure.Code.All
    Dim element     As Variant
    For Each element In LinkedUserforms
        If InStr(1, S, "@INCLUDE USERFORM " & element) = 0 And InStr(1, ListOfImports, "@INCLUDE USERFORM " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE USERFORM " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE USERFORM " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oProcedure.oModule.CodeModule.InsertLines oProcedure.lines.FirstOfBodyAfterComments, ListOfImports
    End If
End Sub

Public Sub InjectLinkedDeclarations()
    Dim ListOfImports As String
    Dim S           As String: S = oProcedure.Code.All
    Dim coll        As New Collection
    Dim element     As Variant
    For Each element In LinkedDeclarations
        If InStr(1, S, "'@INCLUDE DECLARATION " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE DECLARATION " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE DECLARATION " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oProcedure.oModule.CodeModule.InsertLines oProcedure.lines.FirstOfBodyAfterComments, ListOfImports
    End If
End Sub

Public Function LinkedDeclarations() As Collection
    aProject.Initialize(oProcedure.oWorkbook).Declarations.createTable
    Dim TargetWorksheet As Worksheet: Set TargetWorksheet = oProcedure.oWorkbook.Sheets("Declarations_Table")
    Dim coll        As New Collection
    Dim S           As String: S = oProcedure.Code.All
    Dim element
    For Each element In DeclarationsTableKeywords
        If RegexTest(S, "\b ?" & CStr(element) & "\b") Then
            On Error Resume Next
            coll.Add CStr(element), CStr(element)
            On Error GoTo 0
        End If
    Next
    Set LinkedDeclarations = coll
End Function

Function DeclarationsTableKeywords() As Collection
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = oProcedure.oWorkbook.Sheets("Declarations_Table")
    Dim Lr          As Long: Lr = getLastRow(TargetWorksheet)
    Dim coll        As New Collection
    Dim cell        As Range
    For Each cell In TargetWorksheet.Range("C2:C" & Lr)
        On Error Resume Next
        coll.Add cell.text, cell.text
        On Error GoTo 0
    Next
    Set DeclarationsTableKeywords = coll
End Function

Public Function LinkedUserforms()
    Dim coll        As New Collection
    Dim S           As String: S = oProcedure.Code.All
    Dim FormName
    For Each FormName In aModules.Initialize(oProcedure.oWorkbook.VBProject).UserformNames
        If RegexTest(S, "\W" & FormName & "[.(\W]") = True Then coll.Add FormName
    Next
    Set LinkedUserforms = coll
End Function




Public Function LinkedClasses() As Collection
    Dim coll        As New Collection
    Dim var         As Variant
    var = aModule.Initialize(oProcedure.oModule).ClassCalls
    Dim S           As String: S = oProcedure.Code.All
    Dim Keyword     As String
    Dim ClassName   As String
    Dim element     As Variant
    Dim i           As Long
    On Error Resume Next
    For i = LBound(var, 1) To UBound(var, 1)
        If InStr(1, S, var(i, 1)) > 0 Or InStr(1, S, var(i, 2)) > 0 Then
            coll.Add var(i, 1), var(i, 1)
        End If
    Next
    For Each element In aModules.Initialize(oProcedure.oWorkbook.VBProject).ClassNames
        If InStr(1, S, element) > 0 Then
            coll.Add element, CStr(element)
        End If
    Next
    On Error GoTo 0
    Set LinkedClasses = coll
End Function


Function LinkedSheets() As Collection
    Dim SheetNames  As New Collection
    Dim TargetWorksheet As Worksheet
    For Each TargetWorksheet In oProcedure.oWorkbook.Worksheets
        On Error Resume Next
        SheetNames.Add TargetWorksheet.Name
        SheetNames.Add oProcedure.oWorkbook.Worksheets(TargetWorksheet.Name).codeName
        On Error GoTo 0
    Next
    Dim element
    Dim S           As String: S = oProcedure.Code.All
    Dim coll        As New Collection
    For Each element In SheetNames
        If InStr(1, CStr(element), " ") > 0 And InStr(1, S, CStr(element)) > 0 Then
            coll.Add CStr(element)
        ElseIf RegexTest(S, "\W" & CStr(element) & "[.(\W]") = True Then
            coll.Add CStr(element), CStr(element)
        End If
    Next
    Set LinkedSheets = coll
End Function


Public Sub ExportLinkedCode()
    Dim Code        As String: Code = LinkedCode
    Dim MergedName  As String: MergedName = "Merged_" & oProcedure
    Dim fileName    As String: fileName = LOCAL_LIBRARY_PROCEDURES & MergedName & ".txt"
    Debug.Print "OVERWROTE " & MergedName
    TxtOverwrite fileName, Code
    TxtPrependContainedProcedures fileName
    FollowLink fileName
End Sub

Private Function LinkedCode() As String
    Dim MergedString As String: MergedString = oProcedure.Code.All
    Dim Procedure
    For Each Procedure In collLinkedProceduresDeep
        MergedString = MergedString & vbNewLine & aProcedure.Initialize(oProcedure.oWorkbook, , CStr(Procedure)).Code.All
    Next
    LinkedCode = MergedString
End Function

Public Sub ImportDependencies(Optional Overwrite As Boolean)
    ImportProcedureDependencies oProcedure.Name, Overwrite
End Sub

Private Sub ImportProcedureDependencies( _
        Procedure As String, _
        Overwrite As Boolean)

    Dim Module      As VBComponent
    On Error Resume Next
    Set Module = ModuleOfProcedure(oProcedure.oWorkbook, Procedure)
    If Module Is Nothing Then Exit Sub
    On Error GoTo 0
    Dim obj         As String
    Dim Code        As String
    Code = aProcedure.Initialize(oProcedure.oWorkbook, Module, Procedure).Code.All
    Dim var
    var = Split(Code, vbNewLine)
    var = Filter(var, "'@INCLUDE ")
    Dim TextLine    As Variant
    For Each TextLine In var
        TextLine = UCase(Trim(TextLine))
        If TextLine Like "'@INCLUDE *" Then
            obj = Split(TextLine, " ")(2)
            obj = VBA.Replace(obj, vbNewLine, "")
            Select Case True
                Case TextLine Like "'@INCLUDE PROCEDURE *": ImportProcedure obj, Overwrite
                Case TextLine Like "'@INCLUDE CLASS *": ImportClass obj, Overwrite
                Case TextLine Like "'@INCLUDE USERFORM *": ImportUserform obj, Overwrite
                Case TextLine Like "'@INCLUDE DECLARATION *": ImportDeclaration obj
            End Select
        End If
    Next
End Sub

Public Sub Update()
    ImportProcedure oProcedure.Name, True
End Sub

Public Sub ImportProcedure( _
        Procedure As String, _
        Overwrite As Boolean)
    '
    Dim ProcedurePath As String: ProcedurePath = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"
    Dim result      As String
    On Error Resume Next
    result = TxtRead(ProcedurePath)
    On Error GoTo 0

    If Len(result) = 0 Then
        On Error Resume Next
        result = TXTReadFromUrl(GITHUB_LIBRARY_PROCEDURES & Procedure & ".txt")
        On Error GoTo 0
        If Len(result) > 0 And Not UCase(result) Like ("*NOT FOUND*") Then
            TxtOverwrite ProcedurePath, result
        Else
            Debug.Print "File " & Procedure & ".txt not found neither localy nor online"
            Exit Sub
        End If
    End If

    Dim filelastmod: filelastmod = StringLastModified(result)
    Dim proclastmod

    Dim Module      As VBComponent
    If ProcedureExists(oProcedure.oWorkbook, Procedure) = True Then
        Set Module = ModuleOfProcedure(oProcedure.oWorkbook, Procedure)
        proclastmod = ProcedureLastModified(oProcedure.oWorkbook, Module, Procedure)
        If Overwrite = True Then
            If proclastmod = 0 Or proclastmod < filelastmod Then
                ProcedureReplace Module, Procedure, TxtRead(ProcedurePath)
            End If
        End If
    Else
        Set Module = ModuleAddOrSet(oProcedure.oWorkbook, "vbArcImports", vbext_ct_StdModule)
        Module.CodeModule.AddFromFile ProcedurePath
    End If

    ImportProcedureDependencies Procedure, Overwrite
End Sub

Sub ImportDeclaration(DeclarationName As String)
    Dim FilePath    As String
    FilePath = LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt"
    Dim result      As String
    On Error Resume Next
    result = TxtRead(FilePath)
    On Error GoTo 0

    If Len(result) = 0 Then    'CheckPath(filePath) = "I" Then
        On Error Resume Next
        result = TXTReadFromUrl(GITHUB_LIBRARY_DECLARATIONS & DeclarationName & ".txt")
        On Error GoTo 0
        If Len(result) > 0 And Not UCase(result) Like ("*NOT FOUND*") Then
            TxtOverwrite FilePath, result
        Else
            Debug.Print "File " & DeclarationName & ".txt not found localy or online"
            Exit Sub
        End If
    Else

    End If
    If InStr(1, aProject.Initialize(oProcedure.oWorkbook).Code, result, vbTextCompare) > 0 Then Exit Sub
    Dim Module      As VBComponent
    Set Module = ModuleAddOrSet(oProcedure.oWorkbook, "vbArcImports", vbext_ct_StdModule)
    Module.CodeModule.AddFromString FormatVBA7(result)

End Sub

Sub ImportUserform(UserformName As String, _
        Overwrite As Boolean)

    Dim FilePathFrM As String
    FilePathFrM = LOCAL_LIBRARY_USERFORMS & UserformName & ".frm"
    Dim FilePathFrX As String
    FilePathFrX = LOCAL_LIBRARY_USERFORMS & UserformName & ".frx"

    If CheckPath(FilePathFrM) = "I" Then
        On Error Resume Next
        Dim codeFrM As String
        codeFrM = TXTReadFromUrl(GITHUB_LIBRARY_USERFORMS & UserformName & ".frm")
        Dim codeFrX As String
        codeFrX = TXTReadFromUrl(GITHUB_LIBRARY_USERFORMS & UserformName & ".frx")
        On Error GoTo 0
        If Len(codeFrM) > 0 And Len(codeFrX) > 0 Then
            TxtOverwrite FilePathFrM, codeFrM
            TxtOverwrite FilePathFrX, codeFrX
        Else
            Debug.Print "File " & UserformName & ".frm/.frx not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(UserformName, oProcedure.oWorkbook) Then
        If Overwrite = True Then
            oProcedure.oWorkbook.VBProject.VBComponents.Remove oProcedure.oWorkbook.VBProject.VBComponents(UserformName)
        Else
            Exit Sub
        End If
    End If
    oProcedure.oWorkbook.VBProject.VBComponents.Import FilePathFrM
End Sub

Sub ImportClass(ClassName As String, _
        Overwrite As Boolean)

    Dim FilePath    As String
    FilePath = LOCAL_LIBRARY_CLASSES & ClassName & ".cls"
    If CheckPath(FilePath) = "I" Then
        On Error Resume Next
        Dim Code    As String
        Code = TXTReadFromUrl(GITHUB_LIBRARY_CLASSES & ClassName & ".cls")
        On Error GoTo 0
        If Len(Code) > 0 And Not UCase(Code) Like ("*NOT FOUND*") Then
            TxtOverwrite FilePath, Code
        Else
            MsgBox "File " & ClassName & ".cls not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(ClassName, oProcedure.oWorkbook) Then
        If Overwrite = True Then
            oProcedure.oWorkbook.VBProject.VBComponents.Remove oProcedure.oWorkbook.VBProject.VBComponents(ClassName)
        Else
            Exit Sub
        End If
    End If
    oProcedure.oWorkbook.VBProject.VBComponents.Import FilePath
End Sub


Function CallerModules() As Collection
    Dim coll        As New Collection
    Dim Module      As aModule
    Dim matchCollection As New Collection
    For Each Module In aModules.Initialize(aProcedure.oWorkbook.VBProject).Items
        If Module.Contains(oProcedure.Name, True, True, False) Then
            On Error Resume Next
            coll.Add Module, Module.Name 'Module.Name, Module.Name
            On Error GoTo 0
        End If
    Next
    Set CallerModules = coll
End Function

Function CallerModulesToString() As String
    Dim element
    Dim result      As String
    For Each element In CallerModules
        result = result & IIf(result <> "", vbNewLine, "") & element
    Next
    CallerModulesToString = result
End Function

Function Callers() As Collection
    Dim result      As New Collection
    Dim Module      As aModule
    Dim Procedure   As aProcedure
    For Each Module In CallerModules
        For Each Procedure In Module.Procedures
            If Procedure.Name <> oProcedure.Name Then
                If Procedure.Code.Contains(oProcedure.Name, True, True, False) Then
                    On Error Resume Next
                    result.Add Procedure 'Procedure.Name, Procedure.Name
                    Exit For
                    On Error GoTo 0
                End If
            End If
        Next
    Next
    Set Callers = result
End Function

Function CallersToString() As String
    Dim Procedure   As aProcedure
    Dim result      As String
    For Each Procedure In Callers
        result = result & IIf(result <> "", vbNewLine, "") & Procedure.Name
    Next
    CallersToString = result
End Function

Public Sub BringProcedureHere(Procedure As String)
    Dim ap          As New aProcedure
    ap.Initialize , , Procedure
    Dim S           As String
    S = ap.Code.All
    If InStr(1, S, "'@AssignedModule") = 0 Then
        ap.CustomProperties.ParentAssigned = ap.Parent.Name
    End If
    S = ap.Code.All
    ap.Delete
    oProcedure.oModule.CodeModule.InsertLines oProcedure.lines.last + 1, S
End Sub


Sub BringLinkedProceduresHere()
    Dim el
    For Each el In LinkedProceduresDeep
        BringProcedureHere CStr(el)
    Next
End Sub

Sub AddToLinkedTable()

    Dim aw          As aProject
    Set aw = aProject.Initialize(oProcedure.oWorkbook)
    Dim cell        As Range
    Dim TargetWorksheet As Worksheet
    On Error Resume Next
    Set TargetWorksheet = ThisWorkbook.Sheets("Linked_Table_" & aw.NameClean)
    If TargetWorksheet Is Nothing Then Toast "Worksheet Linked_Table_" & aw.NameClean & " does not exist. Create first. Terminating": Exit Sub
    Set cell = TargetWorksheet.Columns(4).Find(oProcedure.Name, LookAt:=xlWhole)
    If Not cell Is Nothing Then Debug.Print "Procedure " & oProcedure.Name & " was already exporterd. Terminating.": Exit Sub
    On Error GoTo 0

    Dim var
    ReDim var(1 To 8)
    var(1) = oProcedure.oWorkbook.Name
    var(2) = aModule.Initialize(oProcedure.oModule).TypeToString
    var(3) = oProcedure.oModule.Name
    var(4) = oProcedure.Name
    var(5) = aCollection.Init(Me.LinkedProcedures).ToString(vbNewLine)
    var(6) = aCollection.Init(Me.LinkedClasses).ToString(vbNewLine)
    var(7) = aCollection.Init(Me.LinkedUserforms).ToString(vbNewLine)
    var(8) = aCollection.Init(Me.LinkedDeclarations).ToString(vbNewLine)

    Dim Lr          As Long
    Lr = getLastRow(TargetWorksheet) + 1
    TargetWorksheet.Range("A" & Lr & ":H" & Lr).value = var
End Sub
