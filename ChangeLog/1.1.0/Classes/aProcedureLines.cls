VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize                As aProcedureLines  :
' Get CountOfBody               As Long             :
' Get CountOfDeclarationLines   As Long             :
' Get CountOfHeaderLines        As Variant          :
' Get FirstOfBody               As Long             :
' Get FirstOfBodyAfterComments  As Long             :
' Get FirstOfDeclaration        As Long             :
' Get FirstOfHeader             As Long             :
' Get LastOfBody                As Long             :
' Get LastOfDeclaration         As Long             :
' Get LastOfHeader              As Long             :
' Get LikeThis                  As Long             :
' Get Longest                   As Long             :
' Get count                     As Long             :
' Get first                     As Long             :
' Get last                      As Long             :
' Sub EnsureBlankLineBeforeProcedure                :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As New aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureLines
    Set oProcedure = Procedure
'    EnsureBlankLineBeforeProcedure ''?'' crashes but works fine on its own
    Set Initialize = Me
End Function

Public Sub EnsureBlankLineBeforeProcedure()
    Dim Line        As Long: Line = FirstOfDeclaration
    If Line = 1 Then oProcedure.Parent.CodeModule.InsertLines 1, "": Exit Sub
    Dim check       As String
    On Error Resume Next
    check = oProcedure.Parent.CodeModule.ProcOfLine(Line - 1, oProcedure.KindAsLong)
    On Error GoTo 0
    If check <> oProcedure.Name And _
    Len(Trim(oProcedure.Parent.CodeModule.lines(first, 1))) > 0 Then _
    oProcedure.Parent.CodeModule.InsertLines Line, ""
End Sub

Public Property Get first() As Long
    first = oProcedure.Parent.CodeModule.procStartLine(oProcedure.Name, oProcedure.KindAsLong)
End Property

Public Property Get last() As Long
    last = first + count - 1
End Property

Public Property Get count() As Long
    count = oProcedure.Parent.CodeModule.ProcCountLines(oProcedure.Name, oProcedure.KindAsLong)
End Property

Public Property Get FirstOfHeader() As Long
    FirstOfHeader = first
End Property
Public Property Get LastOfHeader() As Long
    LastOfHeader = FirstOfDeclaration - 1
End Property
Public Property Get CountOfHeaderLines()
    CountOfHeaderLines = LastOfHeader - FirstOfHeader + 1
End Property

Public Property Get FirstOfDeclaration() As Long
    FirstOfDeclaration = oProcedure.Parent.CodeModule.ProcBodyLine(oProcedure.Name, oProcedure.KindAsLong)
End Property
Public Property Get LastOfDeclaration() As Long
    LastOfDeclaration = FirstOfDeclaration + CountOfDeclarationLines - 1
End Property

Public Property Get CountOfDeclarationLines() As Long
    Dim targetLine  As Long: targetLine = FirstOfDeclaration
    Dim counter     As Long: counter = 1
    Dim s           As String: s = oProcedure.Parent.CodeModule.lines(targetLine, counter)
    Do While Right(s, 1) = "_"
        counter = counter + 1
        s = oProcedure.Parent.CodeModule.lines(targetLine, counter)
    Loop
    CountOfDeclarationLines = counter
End Property

Public Property Get FirstOfBody() As Long
    FirstOfBody = FirstOfDeclaration + CountOfDeclarationLines
End Property
Public Property Get LastOfBody() As Long
    LastOfBody = last - 1
End Property
Public Property Get CountOfBody() As Long
    CountOfBody = LastOfBody - FirstOfBody + 1
End Property
Public Property Get FirstOfBodyAfterComments() As Long
    Dim i           As Long
    Dim s           As String
    For i = FirstOfBody To LastOfBody
        s = Trim(oProcedure.Parent.CodeModule.lines(i, 1))
        If s = vbNullString Then
            Exit For
        ElseIf Left(s, 1) = "'" Then
        ElseIf Left(s, 3) = "Rem" Then
        ElseIf Right(Trim(oProcedure.Parent.CodeModule.lines(i - 1, 1)), 1) = "_" Then
        ElseIf Right(s, 1) = "_" Then
        Else
            Exit For
        End If
    Next
    FirstOfBodyAfterComments = i
End Property

Public Property Get LikeThis(this As String, Optional startAt As Long = 0, Optional endAt As Long = 0) As Long
    If startAt = 0 Then startAt = Me.first
    If endAt = 0 Then endAt = Me.last
    Dim targetLine  As String
    Dim i           As Long
    For i = startAt To endAt
        targetLine = Trim(oProcedure.Parent.CodeModule.lines(i, 1))
        If UCase(targetLine) Like UCase(this) Then
            LikeThis = i
            Exit Property
        End If
    Next
End Property

Public Property Get Longest() As Long
    Longest = LargestLength(Split(oProcedure.code.All, vbNewLine))
End Property
