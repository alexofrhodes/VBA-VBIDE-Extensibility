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
    Dim StartLine As Integer: StartLine = oProcedure.code.lines.first
    oProcedure.Parent.CodeModule.DeleteLines StartLine, oProcedure.code.lines.count
    oProcedure.Parent.CodeModule.InsertLines StartLine, newCode
    oProcedure.Activate
End Sub

Public Sub AddBlankBeforeEachRow()
    Dim arr
    arr = Split(oProcedure.code.All, vbNewLine)
    arr = ArrayRemoveEmptyElements(arr)
    Me.Replace vbNewLine & Join(arr, vbNewLine & vbNewLine)
End Sub

Sub UnfoldDeclaration()
    Dim newCode As String
    newCode = oProcedure.code.DeclarationClean & vbLf & _
             oProcedure.code.Body & vbLf & _
             oProcedure.Parent.CodeModule.lines(oProcedure.code.lines.last, 1)
    Replace newCode
End Sub

Sub FoldDeclaration()

    If oProcedure.arguments.count = 0 Then Exit Sub

    Dim StartLine   As Long
    Dim firstColumn As Long
    StartLine = oProcedure.code.lines.FirstOfDeclaration
    firstColumn = InStr(1, oProcedure.Parent.CodeModule.lines(StartLine, 1), "(")
    
    Dim lastline    As Long
    lastline = oProcedure.code.lines.LastOfDeclaration
    Dim lastColumn As Long
    lastColumn = InStr(1, oProcedure.Parent.CodeModule.lines(lastline, 1), ")") + 1
    If lastColumn = 0 Then Exit Sub
    
    Dim cp As aCodeModule
    Set cp = aCodeModule.Initialize(oProcedure.Parent)
    cp.SetSelection StartLine, firstColumn, lastline, lastColumn
    cp.FoldLine
    cp.SetSelection StartLine, firstColumn, oProcedure.code.lines.LastOfDeclaration, 10000
End Sub


Public Sub CommentsToOwnLine()
    oProcedure.Format.Replace CommentsMoveToOwnLine(oProcedure.code.All)
End Sub


Public Sub Indent()
    Dim s           As String
    Dim nIndent     As Long
    Dim i           As Long
    For i = oProcedure.code.lines.first To oProcedure.code.lines.last
        s = LTrim$(oProcedure.Parent.CodeModule.lines(i, 1))
        If Trim(s) <> "" Then
            If IsBlockEnd(s) Then nIndent = nIndent - 1
            If nIndent < 0 Then nIndent = 0
            s = Space$(nIndent * 4) & s
            oProcedure.Parent.CodeModule.ReplaceLine i, s
            If IsBlockStart(LTrim$(s)) Then nIndent = nIndent + 1
        End If
    Next
End Sub

Public Sub NumbersAdd()
    Dim counter     As Long: counter = 100
    Dim s           As String
    Dim i           As Long
    For i = oProcedure.code.lines.first To oProcedure.code.lines.last
        s = oProcedure.Parent.CodeModule.lines(i, 1)
        If IsLineNumberAble(s) Then
            If Right(Trim(oProcedure.Parent.CodeModule.lines(i - 1, 1)), 1) <> "_" Then
                oProcedure.Parent.CodeModule.ReplaceLine i, counter & ":" & s
                counter = counter + 1
            End If
        Else
            oProcedure.Parent.CodeModule.ReplaceLine i, Space(4) & s
        End If
    Next i
End Sub

Public Sub NumbersRemove()
    Dim s As String
    Dim i As Long
    For i = oProcedure.code.lines.first To oProcedure.code.lines.last
        s = oProcedure.Parent.CodeModule.lines(i, 1)
        If IsNumeric(Left(Trim(s), 1)) Then
            oProcedure.Parent.CodeModule.ReplaceLine i, Mid(s, InStr(s, ":") + 1)
        ElseIf Left(s, 4) = Space(4) Then
            oProcedure.Parent.CodeModule.ReplaceLine i, Mid(s, 5)
        End If
    Next i
End Sub

Public Sub RemoveEmptyLines()
    Dim i As Long
    For i = oProcedure.code.lines.last To oProcedure.code.lines.first Step -1
        If Len(Trim(oProcedure.Parent.CodeModule.lines(i, 1))) = 0 Then
            oProcedure.Parent.CodeModule.DeleteLines i, 1
        End If
    Next
    oProcedure.Parent.CodeModule.InsertLines oProcedure.code.lines.first, ""
    oProcedure.Parent.CodeModule.CodePane.SetSelection oProcedure.code.lines.FirstOfDeclaration, 1, oProcedure.code.lines.FirstOfDeclaration, 1
End Sub


Public Sub CommentsRemove(Optional StartLine As Long = 0, Optional EndLine As Long = 0, Optional RemoveRem As Boolean = True)
        If StartLine = 0 Then StartLine = oProcedure.code.lines.first
        If EndLine = 0 Then EndLine = oProcedure.code.lines.last
        Dim N               As Long
        Dim lineText        As String
        Dim QUOTES          As Long
        Dim Q               As Long
        Dim StartPos        As Long
        Dim l As Long
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
                For l = 1 To N
                    If Mid(lineText, l, 1) = """" Then
                        QUOTES = QUOTES + 1
                    End If
                Next l
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
    Dim s           As String
    Dim i           As Long
    For i = oProcedure.code.lines.last To oProcedure.code.lines.first Step -1
        s = Trim(oProcedure.Parent.CodeModule.lines(i, 1))
        s = UCase(s)
        If s Like this Then
            oProcedure.Parent.CodeModule.DeleteLines i
        End If
    Next
End Sub


Sub BlankLinesToDividers(Optional Character As String = "~")
    Dim l           As Long: l = oProcedure.lines.Longest
    Dim i           As Long
    For i = oProcedure.lines.last To oProcedure.lines.first Step -1
        If Len(Trim(oProcedure.Parent.CodeModule.lines(i, 1))) = 0 Then
            oProcedure.Parent.CodeModule.ReplaceLine i, "'" & String(l - 1, Character)
        End If
    Next
End Sub

Public Sub EnableDebugPrint()
    Dim N           As Long
    Dim s           As String
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            s = .lines(N, 1)
            If Left(Trim(s), 6) = "'Debug" Then
                s = VBA.Replace(s, "'", "", , 1)
                .ReplaceLine N, s
            End If
        Next N
    End With
End Sub
Public Sub DisableDebugPrint()
    Dim N           As Long
    Dim s           As String
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            s = .lines(N, 1)
            If Left(Trim(s), 5) = "Debug." Then
                .ReplaceLine N, "'" & s
            End If
        Next N
    End With
End Sub


Public Sub EnableStop()
    Dim N           As Long
    Dim s           As String
    Dim Keyword     As String
    Keyword = "Stop"
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            s = .lines(N, 1)
            If Not Left(Trim(s), 1) = "'" Then
                If InStrExact(1, s, Keyword) > 0 Then
                    s = VBA.Replace(s, "'", "", , 1)
                    .ReplaceLine N, s
                End If
            End If
        Next N
    End With
End Sub
Public Sub DisableStop()
    Dim N           As Long
    Dim s           As String
    Dim Keyword     As String: Keyword = "Stop"
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            s = .lines(N, 1)
            If Not Left(Trim(s), 1) = "'" Then
                If InStrExact(1, s, Keyword) > 0 Then
                    .ReplaceLine N, "'" & s
                End If
            End If
        Next N
    End With
End Sub

Public Sub CommentsToRem()
    Dim N           As Long
    Dim s           As String
    With oProcedure.Parent.CodeModule
        For N = oProcedure.lines.last To oProcedure.lines.first Step -1
            s = .lines(N, 1)
            If Left(Trim(s), 1) = "'" Then
                .ReplaceLine N, VBA.Replace(s, "'", "Rem ", , 1)
            End If
        Next N
    End With
End Sub
