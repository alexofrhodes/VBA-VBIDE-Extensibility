VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProjectDeclarations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize         As aProjectDeclarations  :
' Fun Items              As Collection            :
' Fun declaredEnums      As String                :
' Fun declaredFunctions  As String                :
' Fun declaredKeywords   As Variant               :
' Fun declaredSubs       As String                :
' Fun declaredTypes      As String                :
' Fun tableKeywords      As Collection            :
' Sub ExportDeclarations                          :
' Sub ExportTable                                 :
' Sub createTable                                 :
'--------------------------------------------------
'@EndFolder ListOfProcedures
Option Explicit

Private oProject As aProject

Public Function Initialize(Project As aProject) As aProjectDeclarations
    Set oProject = Project
    Set Initialize = Me
End Function

Public Function Items( _
        Optional includeScope As Boolean, _
        Optional includeType As Boolean, _
        Optional includeKeywords As Boolean, _
        Optional includeDeclarations As Boolean, _
        Optional includeComponentName As Boolean, _
        Optional includeComponentType As Boolean) As Collection

    Dim ComponentCollection As New Collection
    Dim ComponentTypecollection As New Collection
    Dim DeclarationsCollection As New Collection
    Dim KeywordsCollection As New Collection
    Dim Output      As New Collection
    Dim ScopeCollection As New Collection
    Dim TypeCollection As New Collection

    Dim element     As Variant
    Dim OriginalDeclarations As Variant
    Dim str         As Variant

    Dim tmp         As String
    Dim helper      As String
    Dim i           As Long

    Dim Module      As aModule
    For Each Module In aModules.Initialize(oProject.Project).Items
'        If Module.Type = vbext_ct_StdModule Or Module.Type = vbext_ct_MSForm Then
        If Module.Component.CodeModule.CountOfDeclarationLines > 0 Then
            str = Module.Component.CodeModule.lines(1, Module.Component.CodeModule.CountOfDeclarationLines)
            str = Replace(str, "_" & vbNewLine, "")
            OriginalDeclarations = str
            tmp = str
            Do While InStr(1, str, "End Type") > 0
                tmp = Mid(str, InStr(1, str, "Type "), InStr(1, str, "End Type") - InStr(1, str, "Type ") + 8)
                str = Replace(str, tmp, Split(tmp, vbNewLine)(0))
            Loop
            Do While InStr(1, str, "End Enum") > 0
                tmp = Mid(str, InStr(1, str, "Enum "), InStr(1, str, "End Enum") - InStr(1, str, "Enum ") + 8)
                str = Replace(str, tmp, Split(tmp, vbNewLine)(0))
            Loop
            Do While InStr(1, str, "  ") > 0
                str = Replace(str, "  ", " ")
            Loop

            str = Split(str, vbNewLine)
            tmp = OriginalDeclarations

            For Each element In str
                If Len(CStr(element)) > 0 And Not Trim(CStr(element)) Like "'*" And Not Trim(CStr(element)) Like "Rem*" Then
                    If RegexTest(CStr(element), "\b ?Enum \b") Then
                        KeywordsCollection.Add getKeyword(CStr(element), " ", "Enum")
                        DeclarationsCollection.Add getKeyword(tmp, , "Enum " & KeywordsCollection.Item(KeywordsCollection.count), "End Enum", , , True)
                        TypeCollection.Add "Enum"
                        ComponentCollection.Add Module.Name
                        ComponentTypecollection.Add Module.TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    
                    ElseIf RegexTest(CStr(element), "\b ?Type \b") Then
                        KeywordsCollection.Add getKeyword(CStr(element), " ", "Type")
                        DeclarationsCollection.Add getKeyword(tmp, , "Type " & KeywordsCollection.Item(KeywordsCollection.count), "End Type", , , True)
                        TypeCollection.Add "Type"
                        ComponentCollection.Add Module.Name
                        ComponentTypecollection.Add Module.TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    
                    ElseIf InStr(1, CStr(element), "Const ", vbTextCompare) > 0 Then
                        KeywordsCollection.Add getKeyword(CStr(element), " ", "Const")
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Const"
                        ComponentCollection.Add Module.Name
                        ComponentTypecollection.Add Module.TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    
                    ElseIf RegexTest(CStr(element), "\b ?Sub \b") Then
                        KeywordsCollection.Add getKeyword(CStr(element), " ", "Sub")
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Sub"
                        ComponentCollection.Add Module.Name
                        ComponentTypecollection.Add Module.TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    
                    ElseIf RegexTest(CStr(element), "\b ?Function \b") Then
                        KeywordsCollection.Add getKeyword(CStr(element), " ", "Function")
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Function"
                        ComponentCollection.Add Module.Name
                        ComponentTypecollection.Add Module.TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    
                    ElseIf element Like "*(*) As *" Then
                        helper = Left(element, InStr(1, CStr(element), "(") - 1)
                        helper = Mid(helper, InStrRev(helper, " ") + 1)
                        KeywordsCollection.Add helper
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Other"
                        ComponentCollection.Add Module.Name
                        ComponentTypecollection.Add Module.TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    
                    ElseIf element Like "* As *" Then
                        KeywordsCollection.Add getKeyword(CStr(element), " ", , "As")
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Other"
                        ComponentCollection.Add Module.Name
                        ComponentTypecollection.Add Module.TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    Else
                    End If
                End If
            Next element
        End If
'        End If
    Next Module

    If includeScope = True Then Output.Add ScopeCollection
    If includeType = True Then Output.Add TypeCollection
    If includeKeywords = True Then Output.Add KeywordsCollection
    If includeDeclarations = True Then Output.Add DeclarationsCollection
    If includeComponentType = True Then Output.Add ComponentTypecollection
    If includeComponentName = True Then Output.Add ComponentCollection

    Set Items = Output
End Function

Private Function getKeyword( _
        str As Variant, Optional delim As String _
        , Optional afterWord As String _
        , Optional beforeWord As String _
        , Optional counter As Integer _
        , Optional outer As Boolean _
        , Optional includeWords As Boolean) As String

    Dim i           As Long
    If afterWord = "" And beforeWord = "" And counter = 0 Then
        MsgBox ("Pass at least 1 parameter betweenn -AfterWord- , -BeforeWord- , -counter-")
        Exit Function
    End If
    If TypeName(str) = "String" Then
        If delim <> "" Then
            str = Split(str, delim)
            If UBound(str) <> 0 Then
                If afterWord = "" And beforeWord = "" And counter <> 0 Then
                    If counter - 1 <= UBound(str) Then
                        getKeyword = str(counter - 1)
                        Exit Function
                    End If
                End If
                For i = LBound(str) To UBound(str)
                    If afterWord <> "" And beforeWord = "" Then
                        If i <> 0 Then
                            If str(i - 1) = afterWord Or str(i - 1) = "#" & afterWord Then
                                getKeyword = str(i)
                                Exit Function
                            End If
                        End If
                    ElseIf afterWord = "" And beforeWord <> "" Then
                        If i <> UBound(str) Then
                            If str(i + 1) = beforeWord Or str(i + 1) = "#" & beforeWord Then
                                getKeyword = str(i)
                                Exit Function
                            End If
                        End If
                    ElseIf afterWord <> "" And beforeWord <> "" Then
                        If i <> 0 And i <> UBound(str) Then
                            If (str(i - 1) = afterWord Or str(i - 1) = "#" & afterWord) And (str(i + 1) = beforeWord Or str(i + 1) = "#" & beforeWord) Then
                                getKeyword = str(i)
                                Exit Function
                            End If
                        End If
                    End If
                Next i
            End If
        Else
            If InStr(1, str, afterWord) > 0 And InStr(1, str, beforeWord) > 0 Then
                If includeWords = False Then
                    getKeyword = Mid(str, InStr(1, str, afterWord) + Len(afterWord))
                Else
                    getKeyword = Mid(str, InStr(1, str, afterWord))
                End If
                If outer = True Then
                    If includeWords = False Then
                        getKeyword = Left(getKeyword, InStrRev(getKeyword, beforeWord) - 1)
                    Else
                        getKeyword = Left(getKeyword, InStrRev(getKeyword, beforeWord) + Len(beforeWord) - 1)
                    End If
                Else
                    If includeWords = False Then
                        getKeyword = Left(getKeyword, InStr(1, getKeyword, beforeWord) - 1)
                    Else
                        getKeyword = Left(getKeyword, InStr(1, getKeyword, beforeWord) + Len(beforeWord) - 1)
                    End If
                End If
                Exit Function
            End If
        End If
    Else
        '
    End If
    getKeyword = vbNullString
End Function


Private Function createWorksheet() As Boolean
    If WorksheetExists("Declarations_Table", oProject.WorkbookObject) Then Exit Function
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = oProject.WorkbookObject.Sheets.Add
    With TargetWorksheet
        .Name = "Declarations_Table"
        .Cells.VerticalAlignment = xlVAlignTop
        .Range("A1:F1").value = Split("SCOPE,TYPE,NAME,CODE,MODULE TYPE,MODULE NAME", ",")
        .rows(1).Cells.Font.Bold = True
        .rows(1).Cells.Font.Size = 14
    End With
End Function

Public Sub createTable()

    createWorksheet

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = oProject.WorkbookObject.Sheets("Declarations_Table")
    'if sheet was created within the hour, you probably don't have new declarations
    If Format(Now, "YYMMDDHHNN") - TargetWorksheet.Range("Z1").value < 60 Then Exit Sub

    TargetWorksheet.Range("A2").CurrentRegion.offset(1).clear
    ArrayToRange2D aCollection.CollectionsToArray2D( _
                Items( _
                includeScope:=True, _
                includeType:=True, _
                includeKeywords:=True, _
                includeDeclarations:=True, _
                includeComponentName:=True, _
                includeComponentType:=True)), _
            TargetWorksheet.Range("A2")

    TargetWorksheet.Range("Z1").value = Format(Now, "YYMMDDHHNN")

    sortTable
End Sub

Private Sub sortTable()
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = oProject.WorkbookObject.Worksheets("Declarations_Table")

    Dim sort1       As String: sort1 = "B1"
    Dim sort2       As String: sort2 = "C1"
    Dim sort3       As String    ': sort3 = "D1"

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
        .Header = xlYes
        .Apply
    End With
End Sub






Public Function tableKeywords() As Collection
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = oProject.WorkbookObject.Sheets("Declarations_Table")
    Dim Lr          As Long: Lr = getLastRow(TargetWorksheet)
    Dim coll        As New Collection
    Dim cell        As Range
    For Each cell In TargetWorksheet.Range("C2:C" & Lr)
        On Error Resume Next
        coll.Add cell.text, cell.text
        On Error GoTo 0
    Next
    Set tableKeywords = coll
End Function

Function declaredKeywords() As Variant
    declaredKeywords = WorksheetFunction.Transpose(aCollection.CollectionsToArray2D(Items(, , True)))
End Function

Function declaredEnums() As String
    Dim c           As Variant
    Dim out         As String
    For Each c In aCollection.CollectionsToArray2D(Items(, , , True))
        If InStr(1, CStr(c), "Enum ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    declaredEnums = out
End Function

Function declaredTypes() As String
    Dim c           As Variant
    Dim out         As String
    For Each c In aCollection.CollectionsToArray2D(Items(, , , True))
        If InStr(1, CStr(c), "Type ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    declaredTypes = out
End Function

Function declaredSubs(TargetWorkbook As Workbook) As String
    Dim c           As Variant
    Dim out         As String
    For Each c In aCollection.CollectionsToArray2D(Items(, , , True))
        If InStr(1, CStr(c), "Sub ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    declaredSubs = out
End Function

Function declaredFunctions() As String
    Dim c           As Variant
    Dim out         As String
    For Each c In aCollection.CollectionsToArray2D(Items(, , , True))
        If InStr(1, CStr(c), "Function ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    declaredFunctions = out
End Function





Sub ExportDeclarations(targetFolder As String)
    targetFolder = Replace(targetFolder & "\", "\\", "\")
    FoldersCreate targetFolder

    Dim DeclarationArray As Variant
    DeclarationArray = aCollection.CollectionsToArray2D(Items)
    If TypeName(DeclarationArray) <> "Empty" Then
        TxtOverwrite targetFolder & "Declarations.txt", ArrayToString(DeclarationArray)
    End If
End Sub

Public Sub ExportTable()
    Dim ws          As Worksheet
    Set ws = oProject.WorkbookObject.Sheets("Declarations_Table")
    Dim codeName    As String
    Dim codeText    As String
    Dim cell        As Range
    Dim Lr          As Long
    Lr = getLastRow(ws)
    Dim codeCollection As Collection
    For Each cell In ws.Range("C2:C" & Lr)
        codeName = cell.text
        Set codeCollection = New Collection
        On Error Resume Next
        codeCollection.Add cell.offset(0, 1).text, cell.offset(0, 1).text
        On Error GoTo 0
        Do While cell.offset(1, 0).text = cell.text
            Set cell = cell.offset(1, 0)
            On Error Resume Next
            codeCollection.Add cell.offset(0, 1).text, cell.offset(0, 1).text
            On Error GoTo 0
        Loop
        codeText = aCollection.Init(codeCollection).ToString(IIf(codeCollection.count > 1, vbNewLine, ""))
        TxtOverwrite LOCAL_LIBRARY_DECLARATIONS & codeName & ".txt", codeText
    Next
End Sub

