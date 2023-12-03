VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aCodeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Active                   As aCodeModule      :
' Fun ActiveEnum               As aModuleEnumItem  :
' Fun ActiveType               As aModuleTypeItem  :
' Fun CaseTo                   As aCodeModule      :
' Fun Component                As VBComponent      :
' Fun EncapsulateParenthesis   As aCodeModule      :
' Fun EncapsulateQuotes        As aCodeModule      :
' Fun Initialize               As aCodeModule      :
' Fun Procedure                As aProcedure       :
' Fun ProceduresByDeclaration  As Collection       :
' Fun Region                   As aCodeModule      :
' Fun Regions                  As Variant          :
' Fun SortComma                As aCodeModule      :
' Fun SortLines                As aCodeModule      :
' Fun SortLinesByLength        As aCodeModule      :
' Fun Substitute               As aCodeModule      :
' Fun lines                    As aCodeModule      :
' Get ProceduresByName         As Collection       :
' Sub Align                                        :
' Sub AlignAs                                      :
' Sub AlignColumn                                  :
' Sub AlignComments                                :
' Sub BeautifyFunction                             :
' Sub BringProcedureHere                           :
' Sub Comment                                      :
' Sub CommentTargetLine                            :
' Sub CommentsAddRem                               :
' Sub CommentsRemoveRem                            :
' Sub CommentsToggle                               :
' Sub Copy                                         :
' Sub Cut                                          :
' Sub DeleteSelectedLines                          :
' Sub DimMerge                                     :
' Sub DimSeparate                                  :
' Sub Duplicate                                    :
' Sub Encapsulate                                  :
' Sub EncapsulateMultiple                          :
' Sub FoldLine                                     :
' Sub Increment                                    :
' Sub Inject                                       :
' Sub InjectArgumentStyleFolded                    :
' Sub Insert                                       :
' Sub Move                                         :
' Sub ProcedureActivate                            :
' Sub ProcedureImport                              :
' Sub Rotate                                       :
' Sub RotateCommas                                 :
' Sub RotateEqualInLines                           :
' Sub RotateLines                                  :
' Sub RotateMultiple                               :
' Sub SetSelection                                 :
' Sub Sort                                         :
' Sub Todo                                         :
' Sub UnComment                                    :
' Sub UnFoldLine                                   :
' Sub UncommentTargetLine                          :
' Sub UnremTargetLine                              :
' Sub injectDivider                                :
'--------------------------------------------------
'@EndFolder ListOfProcedures


Option Explicit

Private oWorkbook    As Workbook
Private oModule      As VBComponent
Private cm           As CodeModule

Public firstRow      As Long
Public lastRow       As Long

Public firstColumn   As Long
Public lastColumn    As Long

Public rowCount      As Long

Public TextBeforeSelection      As String
Public SelectedText             As String
Public TextAfterSelection       As String

Public textLength    As Long

Public textOfLines   As String 'without selecting them

Public Enum CaseType
    upper
    lower
    sentence
End Enum

Public Enum MoveType
    Up
    Down
End Enum
'@TODO
'class for aCodeModuleFIND
'where there will be a private type
'to store the coordinates
'and functions like exists as boolean, select as acodemodule etc
'---------------------------------------


Public Function Active() As aCodeModule
    Initialize ActiveModule
    Set Active = Me
End Function

Public Function Initialize(Module As VBComponent) As aCodeModule
    Set oModule = Module
    Set cm = oModule.CodeModule
    Set oWorkbook = WorkbookOfModule(Module)
    AssignVariables
    Set Initialize = Me
End Function

Private Sub AssignVariables()
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    cm.CodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn

    firstColumn = StartColumn
    lastColumn = EndColumn

    firstRow = StartLine
    lastRow = EndLine
    If lastRow > StartLine And EndColumn = 1 Then
        lastRow = EndLine - 1
    End If
    rowCount = lastRow - firstRow + 1

    textOfLines = cm.lines(firstRow, rowCount)

    If lastColumn - firstColumn = 0 Then
        If EndLine > lastRow Then
            lastColumn = Len(cm.lines(lastRow, 1)) + 1
        End If
    End If

    TextAfterSelection = Mid(cm.lines(lastRow, 1), lastColumn)      ' + IIf(Len(SelectedText) > 0, 1, 0))
    TextBeforeSelection = Mid(cm.lines(firstRow, 1), 1, firstColumn - 1)

    If rowCount = 1 And (lastColumn - firstColumn > 0) Then
        textLength = lastColumn - firstColumn
    Else
        textLength = Len(textOfLines) - Len(TextBeforeSelection) - Len(TextAfterSelection) - 1 - Len(CountOfCharacters(textOfLines, vbNewLine)) * 2
    End If

    If lastRow - firstRow = 0 Then
        SelectedText = Mid(cm.lines(firstRow, 1), firstColumn, lastColumn - firstColumn)
        Exit Sub
    End If
    Dim str         As String
    Dim i           As Long
    For i = firstRow To lastRow
        If str = "" Then
            str = Mid(cm.lines(i, 1), firstColumn)
        ElseIf i < lastRow Then
            str = str & vbNewLine & cm.lines(i, 1)
        Else
            str = str & vbNewLine & Left(cm.lines(i, 1), lastColumn - 1)
        End If
    Next
    SelectedText = str

End Sub

Public Function Component() As VBComponent
    Set Component = oModule
End Function

Public Sub SetSelection(StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long)
    cm.CodePane.SetSelection StartLine, StartColumn, EndLine, EndColumn
    AssignVariables
End Sub

Private Sub SelectRows()
    SetSelection firstRow, 1, lastRow, Len(cm.lines(lastRow, 1)) + 1
End Sub

Public Function lines() As aCodeModule
    SelectRows
    Set lines = Me
End Function

Public Sub DeleteSelectedLines()
    cm.DeleteLines firstRow, rowCount
    SetSelection firstRow, 1, firstRow, 1
End Sub

Public Sub Cut()
    SelectRows
    CLIP SelectedText
    Debug.Print SelectedText
    DeleteSelectedLines
End Sub

Public Sub Copy()
    SelectRows
    CLIP textOfLines
    AssignVariables
End Sub

Public Function Duplicate() As aCodeModule
    If rowCount > 1 Then
        SelectRows
    Else
        SelectActualLine
    End If
    cm.InsertLines firstRow + rowCount, SelectedText
    SetSelection firstRow + rowCount, 1, firstRow + rowCount + rowCount - 1, 10000
    Set Duplicate = Me
End Function

Public Function SelectActualLine() As aCodeModule
    Dim RowNumber As Long: RowNumber = firstRow
    Dim limit   As Long: limit = 1
    Dim ap      As New aProcedure: Set ap = aProcedure.Active
    If Not ap Is Nothing Then limit = ap.lines.FirstOfBodyAfterComments
    Do While Len(Trim(cm.lines(RowNumber, 1))) = 0 And RowNumber - 1 >= limit
        RowNumber = RowNumber - 1
    Loop
    If Len(Trim(cm.lines(RowNumber, 1))) = 0 Then Exit Function
    Do While Right(cm.lines(RowNumber - 1, 1), 1) = "_"
        RowNumber = RowNumber - 1
    Loop
    Dim counter As Long: counter = 1
    Do While Right(cm.lines(RowNumber, counter), 1) = "_"
        counter = counter + 1
    Loop
    SetSelection RowNumber, 1, RowNumber + counter - 1, 10000
    Set SelectActualLine = Me
End Function

Public Function Increment(Optional incrementBy As Long = 1, Optional matchIndex As Long = -1) As aCodeModule
    If rowCount > 1 Then
        SelectRows
    Else
        SelectActualLine
    End If
    cm.InsertLines firstRow + rowCount, IncrementNumbersInString(SelectedText, incrementBy, matchIndex)
    SetSelection firstRow + rowCount, 1, firstRow + rowCount + rowCount - 1, 10000
    Set Increment = Me
End Function

Sub Todo()
    Const tFormat = "dd-mm-yyyy hh:nn"
    Dim Module      As VBComponent
    Set Module = oModule
    Dim txtName     As String
    txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")
    Dim nLine       As Long
    nLine = aCodeModule.Initialize(Module).firstRow
    Do While nLine > 1 And Right(Module.CodeModule.lines(nLine, 1), 1) = "_"
        nLine = nLine - 1
    Loop
    Dim sLine       As String
    sLine = Module.CodeModule.lines(nLine, 1)
    Dim sSpec       As String
    sSpec = Space(Len(sLine) - Len(Trim(sLine)))
    Dim sTodo       As String
    sTodo = sSpec & "'* @TODO Created: " & _
            VBA.Format$(Now, tFormat) & _
            " Author: " & txtName & vbCrLf & _
            sSpec & "'* @TODO >"
    Module.CodeModule.InsertLines nLine, vbNewLine & sTodo & vbNewLine
    SetSelection nLine + 2, 1000, nLine + 2, 1000
End Sub

Public Sub Move(Direction As MoveType)

    Dim ap As New aProcedure
    On Error Resume Next
    Set ap = ap.Active
    On Error GoTo 0
    If ap Is Nothing Then Exit Sub
       
    Dim insertAt As Long
    insertAt = firstRow + IIf(Direction = Up, -1, 1)
    
    If Direction = Up And insertAt <= ap.lines.LastOfDeclaration Then Exit Sub
    If Direction = Down And insertAt >= ap.lines.last Then Exit Sub

    cm.DeleteLines firstRow, rowCount
    cm.InsertLines insertAt, textOfLines
    
    SetSelection insertAt, 1, insertAt + rowCount - 1, 10000
End Sub

Public Function CaseTo(this As CaseType) As aCodeModule
    If this = lower Then Inject LCase(SelectedText)
    If this = upper Then Inject UCase(SelectedText)
    If this = sentence Then Inject UCase(SelectedText)
    Set CaseTo = Me
End Function

Public Function ActiveEnum() As aModuleEnumItem
    Dim am As New aModule
    am.Initialize oModule
    Set ActiveEnum = am.Enums.Active
    Set am = Nothing
End Function

Public Function ActiveType() As aModuleTypeItem
    Dim am As New aModule
    am.Initialize oModule
    Set ActiveEnum = am.Types.Active
    Set am = Nothing
End Function

Public Function Substitute(oldValue As String, NewValue As String, Optional count As Long = -1) As aCodeModule
    Dim str         As String: str = SelectedText
    Dim lastRowLength As Long
    lastRowLength = 1 + Len(Replace(Split(str, vbNewLine)(UBound(Split(str, vbNewLine))), oldValue, NewValue))
    Dim code        As String
    code = TextBeforeSelection & _
            Replace(str, oldValue, NewValue, , count, vbTextCompare) & _
            TextAfterSelection
    cm.DeleteLines firstRow, rowCount
    cm.InsertLines firstRow, code
    SetSelection firstRow, _
            firstColumn, _
            firstRow + CountOfCharacters(code, vbNewLine), _
            IIf(CountOfCharacters(code, vbNewLine) > 0, lastRowLength, firstColumn + Len(code))
    Set Substitute = Me
End Function

Public Sub Format_VBA7()
    SelectRows
    Inject FormatVBA7(SelectedText)
End Sub


''''''
'DIMS
''''''


Public Sub DimSeparate()
    Dim coll        As New Collection
    Dim code        As String: code = textOfLines
    Dim firstline   As Long: firstline = firstRow
    Dim iSpace      As Long
    Dim iCount      As Long: iCount = rowCount

    If iCount = 1 Then
        iSpace = IndentationCount(cm.lines(firstline, 1))
        code = cm.lines(firstline, iCount)
        code = Join(Split(code, ", "), vbNewLine & Space(iSpace) & "Dim ")
    Else
        Dim element
        For Each element In Split(code, vbNewLine)
            iSpace = IndentationCount(CStr(element))
            coll.Add Join(Split(element, ", "), vbNewLine & Space(iSpace) & "Dim ")
        Next
        code = aCollection.Init(coll).ToString(vbNewLine)
    End If
    code = StringFormatAlignRowsElements(code, " As ", True)

    cm.DeleteLines firstline, iCount
    cm.InsertLines firstline, code
    SetSelection firstline, 1, firstline + UBound(Split(code, vbNewLine)), 10000
End Sub

Public Sub DimMerge()
    Dim code        As String: code = textOfLines
    Dim firstline   As Long: firstline = firstRow
    Dim iCount      As Long: iCount = rowCount
    Dim iSpace      As Long: iSpace = IndentationCount(cm.lines(firstline, 1))
    code = cm.lines(firstline, iCount)
    If code Like "*" & vbNewLine Then code = Left(code, Len(code) - Len(vbNewLine))
    code = Replace(Join(Split(code, vbNewLine)), "Dim ", ", ")
    code = Replace(code, ",", "Dim", , 1)
    code = Space(iSpace) & RegExpReplace(code, "\s{2,}", " ")
    code = Replace(code, " ,", ",")
    code = Mid(code, 2)
    cm.DeleteLines firstline, iCount
    cm.InsertLines firstline, code
    SetSelection firstline, 1, firstline, 10000
End Sub


''''''''''''''''
'INJECT
''''''''''''''''


Public Sub Inject(str As String)
    Dim code        As String: code = TextBeforeSelection & str & TextAfterSelection
    Dim EndLine     As Long
    'catch error when selected all lines
    If rowCount > cm.countOfLines Then EndLine = EndLine - 1
    cm.DeleteLines firstRow, rowCount
    cm.InsertLines firstRow, code
    If Len(str) > 0 Then
        SetSelection firstRow, _
                firstColumn, _
                firstRow + UBound(Split(str, vbLf)), _
                IIf((UBound(Split(str, vbLf))) = 0, lastColumn, 1 + Len(Split(str, vbLf)(UBound(Split(str, vbLf)))))
                dp 1
    Else
        SetSelection firstRow, firstColumn, firstRow, firstColumn
    End If
End Sub

Public Sub Insert(str As String)
    cm.InsertLines firstRow, str
End Sub

Public Sub injectDivider(Optional Character As String = "~")
    If Len(Trim(cm.lines(firstRow, 1))) = 0 Then
        cm.ReplaceLine firstRow, "'" & String(Procedure.lines.Longest - 1, Character)
    Else
        cm.InsertLines firstRow, "'" & String(Procedure.lines.Longest - 1, Character)
    End If
End Sub

Public Sub InjectArgumentStyleFolded()
    Dim result      As String: result = aProcedure.Initialize(ActiveCodepaneWorkbook, , SelectedText).arguments.MultiLine
    If result <> "" Then Inject result
End Sub

Public Sub BeautifyFunction()
    Dim txt         As String: txt = SelectedText
    If txt = "" Then Exit Sub
    txt = StringFormatFunctionNested(txt)
    IndentForInject txt
    Inject txt
End Sub

Private Sub IndentForInject(ByRef txt As String)
    Dim var: var = Split(txt, vbNewLine): If UBound(var) < 1 Then Exit Sub
    Dim Indent      As Long: Indent = firstColumn: If Indent < 2 Then Exit Sub
    Dim i           As Long
    For i = LBound(var) + 1 To UBound(var)
        var(i) = Space(Indent) & var(i)
        txt = Join(var, vbNewLine)
    Next
End Sub


''''''''''''''''
'TEMPLATES
''''''''''''''''


Private Function CopyTemplateFromSheet(Template As String)
    Dim cell        As Range
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Sheets("Templates")
    On Error Resume Next
    Set cell = ws.Columns(1).SpecialCells(xlCellTypeConstants).Find(Template, LookAt:=xlWhole)
    On Error GoTo 0

    CopyTemplateFromSheet = cell.offset(0, 1)
End Function

Private Sub InjectTemplateFromSheet(Template As String)
    If Len(Selection) > 0 Then Exit Sub
    Dim ws          As Worksheet: Set ws = ThisWorkbook.Sheets("Templates")
    Dim cell        As Range
    On Error Resume Next
    Set cell = ws.Columns(1).SpecialCells(xlCellTypeConstants).Find(Template, LookAt:=xlWhole)
    On Error GoTo 0
    If Not cell Is Nothing Then Inject cell.offset(0, 1)
End Sub


''''''''''''''''
'ENCAPSULATE
''''''''''''''''

Public Function EncapsulateQuotes() As aCodeModule
    Encapsulate Chr(34), Chr(34)
    Set EncapsulateQuotes = Me
End Function

Public Function EncapsulateParenthesis() As aCodeModule
    Encapsulate "(", ")"
    Set EncapsulateParenthesis = Me
End Function

Public Sub Encapsulate(before As String, after As String)
    Dim str         As String: str = SelectedText
    Dim code        As String
    code = TextBeforeSelection & _
            before & str & after & _
            TextAfterSelection
    cm.DeleteLines firstRow, rowCount
    cm.InsertLines firstRow, code
    SetSelection firstRow, firstColumn, lastRow, lastColumn + 2
End Sub

Public Sub EncapsulateMultiple( _
                              before As String, _
                              after As String, _
                              Splitter As String)
    Dim code        As String: code = SelectedText
    Dim arr: arr = Split(code, Splitter)
    Dim counter     As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(Splitter)) = Splitter, Len(Splitter), 0)
        arr(counter) = before & arr(counter) & after
    Next
    code = Join(arr, Splitter)
    code = TextBeforeSelection & code & TextAfterSelection
    cm.DeleteLines firstRow, rowCount
    cm.InsertLines firstRow, code
    SetSelection firstRow, _
                 firstColumn, _
                 lastRow, _
                 lastColumn + ( _
                 (Len(before) + (Len(after))) * (UBound(arr) + 1))
End Sub


''''''''''''''''
'SORT
''''''''''''''''


Public Function SortLines() As aCodeModule
    SelectRows
    Sort vbNewLine, True
    Set SortLines = Me
End Function

Public Function SortLinesByLength() As aCodeModule
    SelectRows
    Dim arr
    arr = Split(SelectedText, vbNewLine)
    ArraySortByLength arr, LBound(arr), UBound(arr)
    Inject Join(arr, vbNewLine)
    Set SortLinesByLength = Me
End Function

Public Function SortComma() As aCodeModule
    Sort ",", True
    Set SortComma = Me
End Function

Public Sub Sort(delimeter As String, caseInsensitive As Boolean)
    Dim code        As String: code = SelectedText
    Dim arr
    arr = Split(code, delimeter)
    arr = SortSelectionArray(arr, caseInsensitive)
    code = Join(arr, delimeter)
    Inject code
End Sub


''''''''''''''''
'ROTATE
''''''''''''''''

Public Sub RotateCommas()
    Rotate ","
End Sub

Public Sub RotateLines()
    Rotate vbNewLine
End Sub

Public Sub Rotate(delim As String)
    Rem Rotate multiple  eg. a,b,c,d -> b,c,d,a
    Dim s           As String
    s = SelectedText
    s = Join(ArrayRotate(Split(s, delim)), delim)
    s = TextBeforeSelection & s & TextAfterSelection
    Dim idx         As Long: idx = firstRow
    cm.DeleteLines idx, rowCount
    cm.InsertLines idx, s
    SetSelection firstRow, firstColumn, lastRow, lastColumn
End Sub

Public Sub RotateEqualInLines()
    RotateMultiple "=", vbNewLine
End Sub

Public Sub RotateMultiple(flipper As String, Splitter As String)
    Dim StartLine   As Long
    Dim StartColumn As Long
    Dim EndLine     As Long
    Dim EndColumn   As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, _
                                                StartColumn, _
                                                EndLine, _
                                                EndColumn
    Dim code        As String: code = SelectedText
    Dim arr         As Variant
    arr = Split(code, Splitter)
    arr = ArrayRemoveEmptyElements(arr)
    Dim counter     As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(Splitter)) = Splitter, Len(Splitter), 0)
        arr(counter) = Split(arr(counter), flipper)(1) & flipper & Split(arr(counter), flipper)(0)
    Next
    code = Join(arr, Splitter)
    code = TextBeforeSelection & _
            code & _
            TextAfterSelection
    cm.DeleteLines StartLine, EndLine - StartLine + 1
    cm.InsertLines StartLine, code
    SetSelection StartLine, StartColumn, EndLine, EndColumn
End Sub


''''''''''''''''
'COMMENTS
''''''''''''''''

Public Sub Comment()
    Dim rowLine     As String
    Dim i           As Long
    For i = firstRow To lastRow
        CommentTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection firstRow, 1, lastRow, 1000
End Sub

Public Sub CommentTargetLine(targetLine As Long)
    Dim sLine       As String: sLine = oModule.CodeModule.lines(targetLine, 1)
    If IsCommentLine(sLine) Then Exit Sub
    Dim dif         As Long: dif = Len(sLine) - Len(LTrim(sLine))
    oModule.CodeModule.ReplaceLine targetLine, Space(dif) & "'" & Trim(sLine)
End Sub

Public Sub UnComment()
    Dim i           As Long
    For i = firstRow To lastRow
        UncommentTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection firstRow, 1, lastRow, 1000
End Sub

Public Sub UncommentTargetLine(targetLine As Long)
    Dim sLine       As String
    With oModule.CodeModule
        sLine = .lines(targetLine, 1)
        If Left(Trim(sLine), 1) = "'" Then
            .ReplaceLine targetLine, Replace(sLine, "'", "", , 1)
        End If
    End With
End Sub

Public Sub CommentsAddRem()
    Dim i           As Long
    For i = firstRow To lastRow
        RemTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection firstRow, 1, lastRow, 1000
End Sub

Private Sub RemTargetLine(targetLine As Long)
    Dim rowLine     As String: rowLine = oModule.CodeModule.lines(targetLine, 1)
    If IsCommentLine(rowLine) Then Exit Sub
    oModule.CodeModule.ReplaceLine targetLine, _
                                   Space( _
                                   Len(rowLine) - Len(LTrim(rowLine))) & "Rem " & Trim(rowLine)
End Sub

Public Sub CommentsRemoveRem()
    Dim i           As Long
    For i = firstRow To lastRow
        UnremTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection firstRow, 1, lastRow, 1000
End Sub

Public Sub UnremTargetLine(targetLine As Long)
    Dim sLine       As String
    With oModule.CodeModule
        sLine = .lines(targetLine, 1)
        If Left(Trim(sLine), 4) = "Rem " Then
            .ReplaceLine targetLine, Replace(sLine, "Rem ", "", , 1)
        End If
    End With
End Sub

Sub CommentsToggle()
    Dim sLine       As String
    Dim i           As Long
    For i = firstRow To lastRow
        sLine = cm.lines(i, 1)
        If IsCommentLine(sLine) Then
            If Trim(sLine) Like "Rem*" Then
                UnremTargetLine i
            ElseIf Trim(sLine) Like "'*" Then
                UncommentTargetLine i
            End If
        Else
            CommentTargetLine i
        End If
    Next
    SetSelection firstRow, 1, lastRow, 10000
End Sub


''''''''''''''''
'ALIGN
''''''''''''''''

Public Sub AlignAs()
    Align " As ", SearchFromLeft:=True
End Sub

Public Sub AlignColumn()
    Align ":", SearchFromLeft:=True
    Align ":", SearchFromLeft:=False
End Sub

Public Sub AlignComments()
    Align "'", SearchFromLeft:=False
End Sub

Public Sub Align( _
                AlignString As String, _
                SearchFromLeft As Boolean, _
                Optional AlignAtColumn As Long)

    Dim s           As String: s = StringFormatAlignRowsElements( _
                                   textOfLines, _
                                   AlignString, _
                                   SearchFromLeft, _
                                   AlignAtColumn)
    Dim LineFirst   As Long: LineFirst = firstRow
    oModule.CodeModule.DeleteLines LineFirst, rowCount
    oModule.CodeModule.InsertLines LineFirst, s
    SelectRows
End Sub


''''''''''''''''
'PROCEDURES
''''''''''''''''


Public Sub ProcedureActivate()
    aProcedure.Initialize(ActiveCodepaneWorkbook, ActiveModule, SelectedText).Activate
End Sub

Public Sub ProcedureImport()
    aProcedure.Active.Dependencies.ImportProcedure SelectedText, False
End Sub

Public Property Get ProceduresByName() As Collection
    Dim out           As New Collection
    Dim s             As String: s = SelectedText
    Dim ProcedureName As String
    Dim Procedures    As Collection
    Set Procedures = aProject.Initialize(oWorkbook).Procedures(True, False, False, False)
    Dim Procedure   As aProcedure
    For Each Procedure In Procedures
        ProcedureName = Procedure.Name
        If RegexTest(s, ProcedureName) Then out.Add ProcedureName, ProcedureName
    Next
    Set ProceduresByName = out
End Property

Public Function ProceduresByDeclaration() As Collection
    Dim out         As New Collection
    Dim s           As String: s = SelectedText
    Dim ProcedureName As String
    Dim Procedures  As Collection
    Set Procedures = ProceduresOfTXT(s)
    Dim Procedure   As aProcedure
    For Each Procedure In Procedures
        ProcedureName = Procedure.Name
        If RegexTest(s, ProcedureName) Then out.Add ProcedureName, ProcedureName
    Next
    Set ProceduresByDeclaration = out
End Function

Public Function Procedure() As aProcedure
    If ActiveProcedure = "" Then
        Toast "The carret is not inside a procedure"
        Exit Function
    End If
    Set Procedure = aProcedure.Active
End Function

Public Sub BringProcedureHere()
    aProcedure.Initialize(ActiveCodepaneWorkbook, ActiveModule, ActiveProcedure).Dependencies.BringProcedureHere SelectedText
End Sub


''''''''''
'OTHER
''''''''''

Function Regions() As Variant

End Function

Function Region() As aCodeModule
    If Len(oModule.CodeModule.lines(firstRow, 1)) = 0 Then Exit Function
    
    Dim Procedure   As New aProcedure:    Set Procedure = aProcedure.Active
    Dim nLine       As Long:              nLine = firstRow
    Dim limit       As Long:              limit = Procedure.lines.LastOfDeclaration
    
    Do While Len(Trim(oModule.CodeModule.lines(nLine - 1, 1))) > 0 _
         And nLine - 1 > limit
         nLine = nLine - 1
    Loop
    Dim StartLine As Long:    StartLine = nLine
    nLine = lastRow
    limit = Procedure.lines.last
    Do While nLine + 1 < limit _
    And Len(Trim(oModule.CodeModule.lines(nLine + 1, 1))) > 0
        nLine = nLine + 1
    Loop
    Dim EndLine As Long: EndLine = nLine
    SetSelection StartLine, 1, EndLine, 10000
    Set Region = Me
End Function

Public Sub UnFoldLine()
    Dim str         As String: str = SelectedText
    If Trim(str) = "" Then Exit Sub
    Dim code        As String
    code = TextBeforeSelection & _
            Join(Split(str, "_" & vbNewLine), " ") & _
            TextAfterSelection

    cm.DeleteLines firstRow, rowCount
    cm.InsertLines firstRow, code
    SetSelection firstRow, firstColumn, firstRow, 10000
End Sub

Public Sub FoldLine()
    If rowCount > 1 Then UnFoldLine
    If InStr(1, SelectedText, ",") = 0 And InStr(1, SelectedText, "(") = 0 Then
        SetSelection firstRow, firstColumn, firstRow, 10000
    End If
    Dim str         As String
    str = SelectedText

    str = Join(Split(str, ", "), ", _" & vbNewLine & Space(firstColumn - 1))
    str = Replace(str, "(", "( _" & vbNewLine & Space(firstColumn - 1), , 1)

    Dim code        As String
    code = TextBeforeSelection & _
            str & _
            TextAfterSelection
    cm.DeleteLines firstRow, rowCount
    cm.InsertLines firstRow, code
    SetSelection firstRow, firstColumn, firstRow + UBound(Split(code, " _" & vbNewLine)), 10000
End Sub
















