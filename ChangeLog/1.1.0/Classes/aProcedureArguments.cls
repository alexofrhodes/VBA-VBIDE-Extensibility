VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureArguments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize  As aProcedureArguments  :
' Fun MultiLine   As String               :
' Fun SingleLine  As String               :
' Get AsSeen      As String               :
' Get Items       As Variant              :
' Get count       As Long                 :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit


Private oProcedure As aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureArguments
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Function Add(Optional argDescription As String, Optional position As Long = -1) As aProcedureArgumentsItem
    If argDescription = "" Then argDescription = "Arg" & Me.count + 1
    If RegexTest(oProcedure.code.DeclarationClean, "\b" & Split(argDescription, " ")(0) & "\b", , True) Then Exit Function
    Dim i As Long
    For i = oProcedure.lines.FirstOfDeclaration To oProcedure.lines.LastOfDeclaration
        Dim s As String
        s = oProcedure.Parent.CodeModule.lines(i, 1)
        If Me.count = 0 And InStr(1, s, "()") > 0 Then
            oProcedure.Parent.CodeModule.ReplaceLine i, VBA.Replace(s, "()", "(" & argDescription & ")", 1, 1)
            GoTo NormalExit
        Else
            Dim it As aProcedureArgumentsItem
            If position = -1 Then
                Set it = oProcedure.arguments.Items(oProcedure.arguments.count)
            ElseIf position = 1 Then
                Set it = oProcedure.arguments.Items(position)
            End If
            If InStr(1, s, it.OriginalString) > 0 Then
                oProcedure.Parent.CodeModule.ReplaceLine i, VBA.Replace(s, it.OriginalString, _
                                                         IIf(position = 1, argDescription & "," & it.OriginalString, _
                                                         it.OriginalString & "," & argDescription), 1, 1)
                GoTo NormalExit
            End If
        End If
    Next
    Exit Function
NormalExit:
    Set Add = oProcedure.arguments.Items(oProcedure.arguments.count)
End Function

Public Property Get AsSeen() As String
    Dim str
    str = oProcedure.code.DeclarationClean
    str = Right(str, Len(str) - InStr(1, str, "("))
    str = Left(str, InStrRev(str, ")") - 1)
    AsSeen = str
End Property

Public Property Get Items(Optional index As Variant)
    Dim coll As Collection
    Set coll = myItems

    Dim this As aProcedureArgumentsItem '<-------------
    If IsMissing(index) Then
        Set Items = coll
    ElseIf IsNumeric(index) Then
        Set Items = coll(index)
    ElseIf TypeName(index) = "String" Then
        Dim argItem As aProcedureArgumentsItem
        For Each this In coll
            If this.Name = index Then
                Set Items = this
                Exit For
            End If
        Next
    End If
End Property

Private Function myItems() As Collection
    Dim coll As Collection
    Set coll = New Collection
    Dim str
    str = AsSeen
    If InStr(1, str, Chr(34) & "," & Chr(34)) > 0 Then
        str = VBA.Replace(str, Chr(34) & "," & Chr(34), Chr(34) & "|" & Chr(34))
    End If
    str = Split(str, ",")
    If UBound(str) = -1 Then GoTo NormalExit
    Dim i As Long
    Dim Output As Variant
    ReDim Output(LBound(str) To UBound(str))
    Dim this As aProcedureArgumentsItem '<-------------
    For i = LBound(str) To UBound(str)
        str(i) = Trim(VBA.Replace(str(i), Chr(34) & "|" & Chr(34), Chr(34) & "," & Chr(34)))
        
        Set this = New aProcedureArgumentsItem
        this.Initialize oProcedure
        
        '@TODO move the following to aProcedureArgumentsItem ???
        this.index = i + 1
        this.OriginalString = str(i)
        this.Name = Split(str(i), " As ")(0)
        this.Name = Split(this.Name, "(")(0)
        this.Name = Mid(this.Name, InStr(1, this.Name, " ") + 1)
        If InStr(1, str(i), " As ") > 0 Then this.IsType = Split(str(i), " As ")(1)
        this.IsOptional = str(i) Like "Optional *"
        this.hasDefaultValue = InStr(1, str(i), "=") > 0
        If InStr(1, str(i), "=") > 0 Then this.DefaultValue = Split(str(i), "=")(1)
        this.IsByVal = InStr(1, str(i), "ByVal ") > 0
        this.IsByRef = Not this.IsByVal
        this.IsParamArray = InStr(1, str(i), "ParamArray ") > 0
        '--------------------------------------------------------
        coll.Add this
    Next
NormalExit:
    Set myItems = coll
End Function

Public Property Get count() As Long
    count = myItems.count 'UBound(Split(oProcedure.Code.DeclarationClean, ",")) + 1
End Property


Public Function MultiLine() As String
    If count = 0 Then MultiLine = oProcedure.Name & "()": Exit Function
    Dim str As Variant, s As String
    Dim firstPart As String, secondPart As String, Output As String
    str = oProcedure.code.DeclarationClean
    Output = oProcedure.Name & "( _"
    Dim indentation As String
    indentation = Space(Len(Output) - 1)
    str = Right(str, Len(str) - InStr(1, str, "("))
    str = Left(str, InStrRev(str, ")") - 1)
    If InStr(1, str, Chr(34) & "," & Chr(34)) > 0 Then
        str = VBA.Replace(str, Chr(34) & "," & Chr(34), Chr(34) & "|" & Chr(34))
    End If
    str = Split(str, ",")
    Dim i           As Long
    For i = LBound(str) To UBound(str)
        str(i) = VBA.Replace(str(i), Chr(34) & "|" & Chr(34), Chr(34) & "," & Chr(34))

    Next
    For i = LBound(str) To UBound(str)
        s = Trim(str(i))

        s = VBA.Replace(s, "Optional ", "")
        s = VBA.Replace(s, "As ", "")
        s = VBA.Replace(s, "ByVal ", "")
        s = VBA.Replace(s, "ByRef ", "")
        s = VBA.Replace(s, "ParamArray ", "")
        s = VBA.Replace(s, "_", "")

        firstPart = Split(s, " ")(0)
        If InStr(1, s, " ") Then
            secondPart = Split(s, " ")(1)
        Else
            secondPart = "Variant"
        End If
        Output = Output & vbNewLine & indentation & firstPart & ":= " & "as" & secondPart & IIf(i <> UBound(str), ", _", ")")
    Next

    MultiLine = Output
End Function


Public Function SingleLine() As String
    Dim var         As Variant
    Dim s           As String
    var = Split(MultiLine, vbNewLine)
    var = ArrayTrim(var)
    If UBound(var) = -1 Then
        SingleLine = oProcedure.Name & "()"
        Exit Function
    End If
    s = Join(var, vbNewLine)
    s = VBA.Replace(s, " _" & vbNewLine, "")
    SingleLine = s
End Function



