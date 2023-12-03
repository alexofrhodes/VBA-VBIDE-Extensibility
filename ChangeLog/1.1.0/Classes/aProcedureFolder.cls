VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Exists      As Boolean           :
' Fun Initialize  As aProcedureFolder  :
' Sub Append                           :
' Sub Create                           :
' Sub Delete                           :
' Sub Overwrite                        :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureFolder
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Sub Overwrite(FolderName As String, this As String)
    Delete FolderName
    Append FolderName, this
End Sub
Public Sub Create(FolderName As String)
    If Me.Exists(FolderName) Then Exit Sub
    If oProcedure.code.ContainsInBody("'@FOLDER ", False, False, False) Then Exit Sub
    Dim StartLine As Long, EndLine As Long
    Dim fline As Long: fline = oProcedure.lines.FirstOfBodyAfterComments
    oProcedure.oModule.CodeModule.InsertLines fline, Join( _
                                                Array( _
                                                    "'@FOLDER " & FolderName, _
                                                    "'" & String(50, "-"), _
                                                    "'" & String(50, "-"), _
                                                    "'@EndFolder " & FolderName), _
                                                vbNewLine)
End Sub

Public Sub Append(FolderName As String, this As String)
    Create FolderName
    oProcedure.oModule.CodeModule.InsertLines oProcedure.code.lines.LikeThis("'@EndFolder " & FolderName) - 1, this
End Sub
Public Function Exists(FolderName As String) As Boolean
    Dim StartLine As Long, EndLine As Long
    Dim sLine       As String
    Dim i           As Long
    For i = oProcedure.lines.FirstOfBody To oProcedure.lines.last
        sLine = Trim(oProcedure.oModule.CodeModule.lines(i, 1))
        If sLine = "'@FOLDER " & FolderName Then
            Exists = True
            Exit Function
        End If
    Next
End Function

Public Sub Delete(FolderName As String)
    If Not Exists(FolderName) Then Exit Sub
    Dim StartLine As Long, EndLine As Long
    Dim sLine       As String
    Dim i           As Long
    For i = oProcedure.lines.FirstOfBody To oProcedure.lines.last
        sLine = Trim(oProcedure.oModule.CodeModule.lines(i, 1))
        If sLine = "'@FOLDER " & FolderName Then StartLine = i
        If StartLine <> 0 Then
            If sLine = "'@EndFolder " & FolderName Then EndLine = i
        End If
        If StartLine <> 0 And EndLine <> 0 Then Exit For
    Next
    If StartLine = 0 Or EndLine = 0 Then Debug.Print "ERROR at aProcedureFolder.Delete": End
    oProcedure.oModule.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
End Sub


