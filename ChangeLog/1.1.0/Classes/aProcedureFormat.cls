VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize  As aProcedureFormat  :
' Sub BlankLinesToDividers             :
' Sub CommentsRemove                   :
' Sub CommentsToOwnLine                :
' Sub CommentsToRem                    :
' Sub DisableDebugPrint                :
' Sub DisableStop                      :
' Sub EnableDebugPrint                 :
' Sub EnableStop                       :
' Sub FoldDeclaration                  :
' Sub Indent                           :
' Sub NumbersAdd                       :
' Sub NumbersRemove                    :
' Sub RemoveEmptyLines                 :
' Sub RemoveLinesLike                  :
' Sub Replace                          :
' Sub UnfoldDeclaration                :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureFormat
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Sub Replace(newCode As String)
    Dim StartLine As Integer: StartLine = oProcedure.Code.lines.first
    oProcedure.Parent.CodeModule.DeleteLines StartLine, oProcedure.Code.lines.count
    oProcedure.Parent.CodeModule.InsertLines StartLine, newCode
End Sub

Sub UnfoldDeclaration()
    Dim newCode As String
    newCode = oProcedure.Code.DeclarationClean & vbLf & _
             oProcedure.Code.Body & vbLf & _
             oProcedure.Parent.CodeModule.lines(oProcedure.Code.lines.last, 1)
    Replace newCode
End Sub

Sub FoldDeclaration()

    If oProcedure.arguments.count = 0 Then Exit Sub

    Dim StartLine   As Long
    Dim firstColumn As Long
    StartLine = oProcedure.Code.lines.FirstOfDeclaration
    firstColumn = InStr(1, oProcedure.Parent.CodeModule.lines(StartLine, 1), "(")
    
    Dim lastline    As Long
    lastline = oProcedure.Code.lines.LastOfDeclaration
    Dim lastColumn As Long
    lastColumn = InStr(1, oProcedure.Parent.CodeModule.lines(lastline, 1), ")") + 1
    If lastColumn = 0 Then Exit Sub
    
    Dim cp As aCodeModule
    Set cp = aCodeModule.Initialize(oProcedure.Parent)
    cp.SetSelection StartLine, firstColumn, lastline, lastColumn
    cp.FoldLine
    cp.SetSelection StartLine, firstColumn, oProcedure.Code.lines.LastOfDeclaration, 10000
End Sub


Public Sub CommentsToOwnLine()
    oProcedure.Format.Replace CommentsMoveToOwnLine(oProcedure.Code.All)
End Sub


Public Sub Indent()
    Dim S           As String
    Dim nIndent     As Long
    Dim i           As Long
    For i = oProcedure.Code.lines.first To oProcedure.Code.lines.last
        S = LTrim$(oProcedure.Parent.CodeModule.lines(i, 1))
        If Trim(S) <> "" Then
            If IsBlockEnd(S) Then nIndent = nIndent - 1
            If nIndent < 0 Then nIndent = 0
            S = Space$(nIndent * 4) & S
            oProcedure.Parent.CodeModule.ReplaceLine i, S
            If IsBlockStart(LTrim$(S)) Then nIndent = nIndent + 1
        End If
    Next
End Sub

Public Sub NumbersAdd()
    Dim counter     As Long: counter = 100
    Dim S           As String
    Dim i           As Long
    For i = oProcedure.Code.lines.first To oProcedure.Code.lines.last
        S = oProcedure.Parent.CodeModule.lines(i, 1)
        If IsLineNumberAble(S) Then
            If Right(Trim(oProcedure.Parent.CodeModule.lines(i - 1, 1)), 1) <> "_" Then
                oProcedure.Parent.CodeModule.ReplaceLine i, counter & ":" & S
                counter = counter + 1
            End If
        Else
            oProcedure.Parent.CodeModule.ReplaceLine i, Space(4) & S
        End If
    Next i
End Sub

Public Sub NumbersRemove()
    Dim S As String
    Dim i As Long
    For i = oProcedure.Code.lines.first To oProcedure.Code.lines.last Step -1
        S = oProcedure.Parent.CodeModule.lines(i, 1)
        If IsNumeric(Left(Trim(S), 1)) Then
            oProcedure.Parent.CodeModule.ReplaceLine i, Mid(S, InStr(S, ":") + 1)
        ElseIf Left(S, 4) = Space(4) Then
            oProcedure.Parent.CodeModule.ReplaceLine i, Mid(S, 5)
        End If
    Next i
End Sub

Public Sub RemoveEmptyLines()
    Dim i As Long
    For i = oProcedure.Code.lines.last To oProcedure.Code.lines.first Step -1
        If Len(Trim(oProcedure.Parent.CodeModule.lines(i, 1))) = 0 Then
            oProcedure.Parent.CodeModule.DeleteLines i, 1
        End If
    Next
End Sub


Public Sub CommentsRemove(Optional StartLine As Long = 0, Optional EndLine As Long = 0, Optional RemoveRem As Boolean = True)
        If StartLine = 0 Then StartLine = oProcedure.Code.lines.first
        If EndLine = 0 Then EndLine = oProcedure.Code.lines.last
        Dim N               As Long
        Dim lineText        As String
        Dim QUOTES          As Long
        Dim Q               As Long
        Dim StartPos        As Long
        Dim L As Long
        Dim j As Long
        For j = EndLine To StartLine Step -1
            lineText = LTrim(oProcedure.Parent.CodeModule.lines(j, 1))
            If Not RemoveRem Then If lineText Like "Rem *" Then GoTo SKIP
            StartPos = 1
Retry:
            N = InStr(StartPos, lineText, "'")
            Q = InStr(StartPos, lineText, """")
            QUOTES = 0
            If Q < N Then
                For L = 1 To N
                    If Mid(lineText, L, 1) = """" Then
                        QUOTES = QUOTES + 1
                    End If
                Next L
            End If
            If QUOTES = Application.WorksheetFunction.Odd(QUOTES) Then
                StartPos = N + 1
                GoTo Retry:
            Else
                Select Case N
                    Case Is = 0
    
                    Case Is = 1
                        oProcedure.Parent.CodeModule.DeleteLines j, 1
                    Case Is > 1
                        oProcedure.Parent.CodeModule.ReplaceLine j, Left(oProcedure.Parent.CodeModule.lines(j, 1), N - 1)
                End Select
            End If
SKIP:
        Next j
End Sub

Public Sub RemoveLinesLike(this As String)
    this = UCase(this)
    Dim S           As String
    Dim i           As Long
    For i = oProcedure.Code.lines.last To oProcedure.Code.lines.first Step -1
        S = Trim(oProcedure.Parent.CodeModule.lines(i, 1))
        S = UCase(S)
        If S Like this Then
            oProcedure.Parent.CodeModule.DeleteLines i
        End If
    Next
End Sub


Sub BlankLinesToDividers(Optional Character As String = "~")
    Dim L           As Long: L = oProcedure.lines.Longest
    Dim i           As Long
    For i = oProcedure.lines.last To oProcedure.lines.first Step -1
        If Len(Trim(oProcedure.Parent.CodeModule.lines(i, 1))) = 0 Then
            oProcedure.Parent.CodeModule.ReplaceLine i, "'" & String(L - 1, Character)
        End If
    Next
End Sub

Public Sub EnableDebugPrint()
    Dim N           As Long
    Dim S           As String
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            S = .lines(N, 1)
            If Left(Trim(S), 6) = "'Debug" Then
                S = VBA.Replace(S, "'", "", , 1)
                .ReplaceLine N, S
            End If
        Next N
    End With
End Sub
Public Sub DisableDebugPrint()
    Dim N           As Long
    Dim S           As String
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            S = .lines(N, 1)
            If Left(Trim(S), 5) = "Debug." Then
                .ReplaceLine N, "'" & S
            End If
        Next N
    End With
End Sub


Public Sub EnableStop()
    Dim N           As Long
    Dim S           As String
    Dim Keyword     As String
    Keyword = "Stop"
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            S = .lines(N, 1)
            If Not Left(Trim(S), 1) = "'" Then
                If InStrExact(1, S, Keyword) > 0 Then
                    S = VBA.Replace(S, "'", "", , 1)
                    .ReplaceLine N, S
                End If
            End If
        Next N
    End With
End Sub
Public Sub DisableStop()
    Dim N           As Long
    Dim S           As String
    Dim Keyword     As String: Keyword = "Stop"
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            S = .lines(N, 1)
            If Not Left(Trim(S), 1) = "'" Then
                If InStrExact(1, S, Keyword) > 0 Then
                    .ReplaceLine N, "'" & S
                End If
            End If
        Next N
    End With
End Sub

Public Sub CommentsToRem()
    Dim N           As Long
    Dim S           As String
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            S = .lines(N, 1)
            If Left(Trim(S), 1) = "'" Then
                .ReplaceLine N, VBA.Replace(S, "'", "Rem ", , 1)
            End If
        Next N
    End With
End Sub
