VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureCustomProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize      As aProcedureCustomProperties  :
' Get Ignore          As Boolean                     :
' Get LastModified    As String                      :
' Get ParentAssigned  As String                      :
' Let Ignore                                         :
' Let LastModified                                   :
' Let ParentAssigned                                 :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Private Type tAprocedureCustomProperties
    LastModified As String
    ParentAssigned As String
    Ignore As Boolean
End Type

Private this As tAprocedureCustomProperties

Public Function Initialize(Procedure As aProcedure) As aProcedureCustomProperties
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Property Get LastModified() As String
    Dim targetLine  As String
    targetLine = oProcedure.lines.LikeThis("*'@LastModified *", oProcedure.lines.FirstOfBody, oProcedure.lines.LastOfBody)
    If targetLine = 0 Then Exit Property
    this.LastModified = Split(Trim(oProcedure.oModule.CodeModule.lines(targetLine, 1)), " ")(1)
    LastModified = this.LastModified
End Property

Public Property Let LastModified(ByVal RHS As String)
    Dim targetLine  As String
    targetLine = oProcedure.lines.LikeThis("*'@LastModified *", oProcedure.lines.FirstOfBody, oProcedure.lines.LastOfBody)
    If targetLine = 0 Then
        oProcedure.Inject.BodyTop "'@LastModified " & RHS, True
    Else
        Dim S As String: S = Trim(oProcedure.oModule.CodeModule.lines(targetLine, 1))
        oProcedure.oModule.CodeModule.ReplaceLine targetLine, Split(S)(0) & " " & RHS
    End If
    this.LastModified = RHS
End Property

Public Property Get ParentAssigned() As String
    Dim targetLine  As String
    targetLine = oProcedure.lines.LikeThis("*'@ParentAssigned *", oProcedure.lines.FirstOfBody, oProcedure.lines.LastOfBody)
    If targetLine = 0 Then Exit Property
    this.ParentAssigned = Split(Trim(oProcedure.oModule.CodeModule.lines(targetLine, 1)), " ")(1)
    ParentAssigned = this.ParentAssigned
End Property

Public Property Let ParentAssigned(ByVal RHS As String)
    Dim targetLine  As String
    targetLine = oProcedure.lines.LikeThis("*'@ParentAssigned *", oProcedure.lines.FirstOfBody, oProcedure.lines.LastOfBody)
    If targetLine = 0 Then
        oProcedure.Inject.BodyTop "'@ParentAssigned " & RHS, True
    Else
        Dim S As String: S = Trim(oProcedure.oModule.CodeModule.lines(targetLine, 1))
        oProcedure.oModule.CodeModule.ReplaceLine targetLine, Split(S)(0) & " " & RHS
    End If
    this.ParentAssigned = RHS
End Property

Public Property Get Ignore() As Boolean
    this.Ignore = oProcedure.code.lines.LikeThis("'@IGNORE", oProcedure.lines.FirstOfBody, oProcedure.lines.LastOfBody) > 0
    Ignore = this.Ignore
End Property

Public Property Let Ignore(ByVal RHS As Boolean)
    Dim targetLine  As Long
    targetLine = oProcedure.lines.LikeThis("*'@IGNORE", oProcedure.lines.FirstOfBody, oProcedure.lines.LastOfBody)
    If targetLine > 0 And RHS = False Then
        oProcedure.oModule.CodeModule.DeleteLines targetLine, 1
    ElseIf targetLine = 0 And RHS = True Then
        oProcedure.Inject.BodyTop "'@IGNORE", False
    End If
    this.Ignore = RHS
End Property



