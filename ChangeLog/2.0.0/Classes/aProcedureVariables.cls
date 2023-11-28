VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize  As aProcedureVariables  :
' Get Items       As Variant              :
' Get count       As Long                 :
' Sub ToImmediate                         :
' Sub UpdatableVariableAdd                :
' Sub UpdatableVariableRemove             :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureVariables
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Sub ToImmediate()
    Dim vars As Collection: Set vars = myItems
    If vars.count = 0 Then Debug.Print "0 variables local to this procedure"
    Dim arr
    ReDim arr(0 To vars.count, 1 To 4)
    Dim i As Long
    arr(0, 1) = "LINE"
    arr(0, 2) = "NAME"
    arr(0, 3) = "TYPE"
    arr(0, 4) = "ASSIGNED?"
    For i = 1 To vars.count
        arr(i, 1) = vars(i).Line
        arr(i, 2) = vars(i).Name
        arr(i, 3) = vars(i).IsType
        arr(i, 4) = vars(i).isAssignedValue
    Next
    dp PrettyPrint.ArrayToTable(arr, True)
End Sub

Public Property Get count() As Long
    count = myItems.count
End Property

Public Property Get Items(Optional index As Variant)
    Dim coll As Collection
    Set coll = myItems
    If coll.count = 0 Then Exit Property
    Dim this As aProcedureVariablesItem '<-------
    If IsMissing(index) Then
        Set Items = coll
    ElseIf IsNumeric(index) Then
        Set Items = coll(index)
    ElseIf TypeName(index) = "String" Then
        For Each this In coll
            If this.Name = index Then
                Set Items = this
                Exit For
            End If
        Next
    Else
    
    End If
End Property

Private Function myItems()
    Dim output As Collection
    Set output = New Collection
    Dim codeLines As Variant: codeLines = Split(oProcedure.code.All, vbNewLine)
    Dim tmp: ReDim tmp(0 To 0)

    Dim i As Long
    For i = LBound(codeLines) To UBound(codeLines)
        codeLines(i) = Trim(codeLines(i))
        If codeLines(i) Like "*:*" Then codeLines(i) = Split(codeLines(i), ":")(0)
    Next

    codeLines = ArrayFilterLike(codeLines, "Dim *", True)
    If UBound(codeLines) = -1 Then Exit Function
    
    Dim element
    For i = LBound(codeLines) To UBound(codeLines)
        codeLines(i) = Trim(StringCommentsRemove(codeLines(i), False))
        If codeLines(i) Like "Dim *,* As *" Then
            For Each element In Split(codeLines(i), ", ")
                If InStr(1, element, "Dim ") = 0 Then element = "Dim " & element
                If InStr(1, element, " As ") = 0 Then element = element & " As "
                tmp(UBound(tmp)) = element
                ReDim Preserve tmp(0 To UBound(tmp) + 1)
            Next
        Else
            For Each element In Split(codeLines(i), ", ")
                tmp(UBound(tmp)) = element
                ReDim Preserve tmp(0 To UBound(tmp) + 1)
            Next
        End If
    Next

    codeLines = ArrayFilterLike(tmp, "Dim *", True)

    Dim this As aProcedureVariablesItem

    For i = LBound(codeLines) To UBound(codeLines)
        Set this = New aProcedureVariablesItem
        this.Initialize oProcedure
        this.Line = oProcedure.lines.LikeThis("*" & CStr(codeLines(i)) & "*")
'        To get the procedure's codeline instead of the module's codeline use
'        this.Line = this.Line - oProcedure.Lines.First + 1
'        or for bodylines oProcedure.Lines.firstofbody
         this.Line = this.Line - oProcedure.lines.FirstOfBody + 1

        If codeLines(i) Like "Dim * As *" Then
            this.Name = Mid(codeLines(i), 5)
            this.Name = Split(this.Name, " As ")(0)
            this.IsType = Split(codeLines(i), " As ")(1)
            If this.IsType Like "New *" Then this.IsType = Split(this.IsType, " ")(1)
        ElseIf codeLines(i) Like "Dim *" Then
            Select Case Right(element, 1)
                Case "$": this.IsType = "String"
                Case "%": this.IsType = "Integer"
                Case "&": this.IsType = "Long"
                Case "!": this.IsType = "Single"
                Case "#": this.IsType = "Double"
                Case "@": this.IsType = "Currency"
                Case "^": this.IsType = "LongLong"
                Case Else: this.IsType = "Variant"
            End Select
            If this.IsType = "Variant" Then
                this.Name = Mid(codeLines(i), 5)
            Else
                this.Name = Mid(codeLines(i), 5, Len(element) - 1)
            End If
        End If
        output.Add this
    Next
    Set myItems = output
End Function

Public Sub UpdatableVariableAdd(this As String)
    Dim i           As Long: i = oProcedure.lines.LikeThis("*UpdatableVariable = *", oProcedure.lines.FirstOfBody, oProcedure.lines.LastOfBody)
    If i = 0 Then Exit Sub
    Dim Q           As String: Q = Chr(34)
    Dim Line        As String: Line = oProcedure.oModule.CodeModule.lines(i, 1)
    If UCase(Trim(Line)) Like UCase("UpdatableVariable = *") And InStr(1, Line, this, vbTextCompare) = 0 Then
        oProcedure.oModule.CodeModule.ReplaceLine i, VBA.Replace(VBA.Replace(Line, Q, Q & this & ",", , 1, vbTextCompare), "," & Q, Q)
        Line = oProcedure.oModule.CodeModule.lines(i, 1)
        If InStr(1, Line, Q & ",") > 0 Then oProcedure.oModule.CodeModule.ReplaceLine i, VBA.Replace(Line, Q & ",", Q)
    End If
End Sub

Public Sub UpdatableVariableRemove(this As String)
    Dim i           As Long: i = oProcedure.lines.LikeThis("*UpdatableVariable = *", oProcedure.lines.FirstOfBody, oProcedure.lines.LastOfBody)
    If i = 0 Then Exit Sub
    Dim Q           As String: Q = Chr(34)
    Dim result      As String
    Dim Line        As String: Line = oProcedure.oModule.CodeModule.lines(i, 1)
    If Trim(Line) Like "UpdatableVariable = *" Then
        If InStr(1, Line, Q & this & Q, vbTextCompare) > 0 Then
            result = VBA.Replace(Line, Q & this & Q, "")
            result = VBA.Replace(result, ",,", ",")
            result = VBA.Replace(result, Q & ",", Q)
            oProcedure.oModule.CodeModule.ReplaceLine i, result
        End If
    End If
End Sub


