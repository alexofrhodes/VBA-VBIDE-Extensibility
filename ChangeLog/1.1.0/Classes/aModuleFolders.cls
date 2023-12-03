VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aModuleFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Exists      As Variant         :
' Fun Initialize  As aModuleFolders  :
' Fun ToString    As String          :
' Sub Append                         :
' Sub Create                         :
' Sub Delete                         :
' Sub Overwrite                      :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oModule As aModule

Public Function Initialize(Module As aModule) As aModuleFolders
    Set oModule = Module
    Set Initialize = Me
End Function

Public Sub Create(FolderName As String)
    If FolderExists(FolderName) Then Exit Sub
    If oModule.HeaderContains("'@FOLDER ") Then Exit Sub

    oModule.Component.CodeModule.InsertLines 1, Join( _
                                                Array( _
                                                    "'@FOLDER " & FolderName, _
                                                    "'" & String(50, "-"), _
                                                    "'" & String(50, "-"), _
                                                    "'@EndFolder " & FolderName), _
                                                vbNewLine)
End Sub

Public Sub Append(FolderName As String, this As String)
    Create FolderName
    Dim fline As Long: fline = oModule.Component.CodeModule.CountOfDeclarationLines
    If fline = 0 Then fline = 1
    Dim sLine       As String
    Dim i As Long
    For i = 1 To fline
        sLine = Trim(oModule.Component.CodeModule.lines(i, 1))
        If sLine = "'@EndFolder " & FolderName Then
            oModule.Component.CodeModule.InsertLines i - 1, this
            Exit Sub
        End If
    Next
End Sub
Public Sub Overwrite(FolderName As String, this As String)
    Delete FolderName
    Append FolderName, this
End Sub
Public Function Exists(FolderName As String)
    Dim StartLine As Long, EndLine As Long
    Dim sLine       As String
    Dim fline As Long: fline = oModule.Component.CodeModule.CountOfDeclarationLines
    If fline = 0 Then Exit Function
    Dim i           As Long
    For i = 1 To fline
        sLine = Trim(oModule.Component.CodeModule.lines(i, 1))
        If sLine = "'@FOLDER " & FolderName Then
            Exists = True
            Exit Function
        End If
    Next
End Function

Public Sub Delete(FolderName As String)
    Dim StartLine As Long, EndLine As Long
    Dim sLine       As String
    Dim fline As Long: fline = oModule.Component.CodeModule.CountOfDeclarationLines
    If fline = 0 Then Exit Sub
    Dim i           As Long
    For i = 1 To fline
        sLine = Trim(oModule.Component.CodeModule.lines(i, 1))
        If sLine = "'@FOLDER " & FolderName Then StartLine = i
        If StartLine <> 0 Then
            If sLine = "'@EndFolder " & FolderName Then EndLine = i
        End If
        If StartLine <> 0 And EndLine <> 0 Then Exit For
    Next

    If StartLine = 0 Or EndLine = 0 Then Exit Sub
    oModule.Component.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
End Sub

Public Function ToString(FolderName As String) As String
    If Not Exists(FolderName) Then Exit Function
    Dim i As Long
    Dim thisLine As String
    Dim out As String
    Dim countOfLines As Long: countOfLines = oModule.Component.CodeModule.CountOfDeclarationLines
    Dim cm As CodeModule: Set cm = oModule.Component.CodeModule
    For i = 1 To countOfLines
        thisLine = cm.lines(i, 1)
        If thisLine = "'@FOLDER " & FolderName Then
            out = thisLine
        ElseIf out <> "" Then
            out = out & vbNewLine & thisLine
            If thisLine = "'@EndFolder " & FolderName Then Exit For
        End If
    Next
    ToString = out
End Function

Public Function countOfLines(FolderName As String) As Long
    If Not Exists(FolderName) Then Exit Function
    Dim i As Long
    Dim thisLine As String
    Dim counter As Long
    Dim LineCount As Long: LineCount = oModule.Component.CodeModule.CountOfDeclarationLines
    Dim cm As CodeModule: Set cm = oModule.Component.CodeModule
    For i = 1 To LineCount
        thisLine = cm.lines(i, 1)
        If thisLine = "'@FOLDER " & FolderName Then
            counter = counter + 1
        ElseIf counter > 0 Then
            counter = counter + 1
            If thisLine = "'@EndFolder " & FolderName Then Exit For
        End If
    Next
    countOfLines = counter
End Function
