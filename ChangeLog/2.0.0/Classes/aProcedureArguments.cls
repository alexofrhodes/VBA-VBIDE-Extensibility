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
    If UBound(str) = -1 Then
        Set Items = coll
        Exit Function
    End If
    Dim i As Long
    Dim output As Variant
    ReDim output(LBound(str) To UBound(str))
    Dim this As aProcedureArgumentsItem '<-------------
    For i = LBound(str) To UBound(str)
        str(i) = Trim(VBA.Replace(str(i), Chr(34) & "|" & Chr(34), Chr(34) & "," & Chr(34)))
        
        Set this = New aProcedureArgumentsItem
        this.Initialize oProcedure
        
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
        coll.Add this
    Next
End Function

Public Property Get count() As Long
    count = myItems.count 'UBound(Split(oProcedure.Code.DeclarationClean, ",")) + 1
End Property


Public Function MultiLine() As String
    If count = 0 Then Exit Function
    Dim str As Variant, S As String
    Dim firstPart As String, secondPart As String, output As String
    str = oProcedure.code.DeclarationClean
    output = oProcedure.Name & "( _"
    Dim indentation As String
    indentation = Space(Len(output) - 1)
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
        S = Trim(str(i))

        S = VBA.Replace(S, "Optional ", "")
        S = VBA.Replace(S, "As ", "")
        S = VBA.Replace(S, "ByVal ", "")
        S = VBA.Replace(S, "ByRef ", "")
        S = VBA.Replace(S, "ParamArray ", "")
        S = VBA.Replace(S, "_", "")

        firstPart = Split(S, " ")(0)
        If InStr(1, S, " ") Then
            secondPart = Split(S, " ")(1)
        Else
            secondPart = "Variant"
        End If
        output = output & vbNewLine & indentation & firstPart & ":= " & "as" & secondPart & IIf(i <> UBound(str), ", _", ")")
    Next

    MultiLine = output
End Function

Public Function SingleLine() As String
    Dim var         As Variant
    Dim S           As String
    var = Split(MultiLine, vbNewLine)
    var = ArrayTrim(var)
    If UBound(var) = -1 Then
        SingleLine = oProcedure.Name & "()"
        Exit Function
    End If
    S = Join(var, vbNewLine)
    S = VBA.Replace(S, " _" & vbNewLine, "")
    SingleLine = S
End Function



