VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize     As aProcedureMove  :
' Get IndexInModule  As Long            :
' Sub Bottom                            :
' Sub Copy                              :
' Sub Down                              :
' Sub ToModule                          :
' Sub ToModuleAssigned                  :
' Sub Top                               :
' Sub Up                                :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As New aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureMove
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Property Get IndexInModule() As Long
    Dim Procedures  As New Collection
    Set Procedures = aModule.Initialize(oProcedure.Parent).Procedures.Names
    Dim index       As Long
    index = aCollection.Init(Procedures).IndexOf(oProcedure.Name) '@TODO what about let/get/set?
    IndexInModule = index
End Property

Public Sub Up()
    Dim Procedures      As New Collection:  Set Procedures = aModule.Initialize(oProcedure.Parent).Procedures
    Dim idx             As Long:            idx = IndexInModule
    If idx = 1 Then Exit Sub
    Dim S               As String:          S = oProcedure.Code.All
    Dim TargetProcedure As New aProcedure:  Set TargetProcedure = Procedures(idx - 1)

    oProcedure.Delete
    oProcedure.Parent.CodeModule.InsertLines TargetProcedure.Code.lines.first, S
    idx = oProcedure.Code.lines.FirstOfDeclaration
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1
End Sub

Public Sub Top()
    Dim Procedures      As New Collection: Set Procedures = aModule.Initialize(oProcedure.Parent).Procedures
    Dim idx             As Long:           idx = IndexInModule:
    If idx = 1 Then Exit Sub
    Dim S               As String:         S = oProcedure.Code.All
    Dim TargetProcedure As New aProcedure: Set TargetProcedure = Procedures(1)
    
    oProcedure.Delete
    oProcedure.Parent.CodeModule.InsertLines TargetProcedure.Code.lines.first, S
    idx = oProcedure.Code.lines.FirstOfDeclaration
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1
End Sub

Public Sub Down()
    Dim Procedures      As New Collection:  Set Procedures = aModule.Initialize(oProcedure.Parent).Procedures
    Dim idx             As Long:            idx = IndexInModule
    If idx = Procedures.count Then Exit Sub
    Dim S               As String:          S = oProcedure.Code.All
    Dim TargetProcedure As New aProcedure:  Set TargetProcedure = Procedures(idx + 1)

    oProcedure.Delete
    oProcedure.Parent.CodeModule.InsertLines TargetProcedure.Code.lines.last + 1, S
    idx = oProcedure.Code.lines.FirstOfDeclaration
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1
End Sub

Public Sub Bottom()
    Dim Procedures      As New Collection:  Set Procedures = aModule.Initialize(oProcedure.Parent).Procedures
    Dim idx             As Long:            idx = IndexInModule:
    If idx = Procedures.count Then Exit Sub
    Dim S               As String:          S = oProcedure.Code.All
    Dim TargetProcedure As New aProcedure:  Set TargetProcedure = Procedures(Procedures.count)

    oProcedure.Delete
    oProcedure.Parent.CodeModule.InsertLines TargetProcedure.Code.lines.last + 1, S
    idx = oProcedure.Code.lines.FirstOfDeclaration
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1
End Sub

Public Sub ToModuleAssigned()
    Dim Module As VBComponent
    Set Module = oProcedure.oWorkbook.VBComponents(oProcedure.CustomProperties.ParentAssigned)
    If Module Is Nothing Then Exit Sub
    If Module.Name = oProcedure.Parent.Name Then Exit Sub
    ToModule Module
End Sub

Public Sub ToModule(Module As VBComponent)
    Dim S As String: S = oProcedure.Code.All
    oProcedure.Delete
    Module.CodeModule.InsertLines Module.CodeModule.countoflines + 1, S
    Set oProcedure.oModule = Module
End Sub



Public Sub Copy( _
        TargetWorkbook As Workbook, _
        Overwrite As Boolean)
    Dim Module      As VBComponent
    Dim S           As String
    Dim Procedure   As Variant
    Dim ap          As aProcedure
    For Each Procedure In oProcedure.Dependencies.LinkedProceduresDeep
        Set ap = New aProcedure
        ap.Initialize oProcedure.oWorkbook, , CStr(Procedure)
        S = ap.Code.All
        If Not ProcedureExists(TargetWorkbook, CStr(Procedure)) Then
            Set Module = ModuleAddOrSet(TargetWorkbook, "vbArcImports", vbext_ct_StdModule)
            Module.CodeModule.AddFromString S
        Else
            Set Module = ModuleOfProcedure(TargetWorkbook, CStr(Procedure))
            If Overwrite = True Then ProcedureReplace Module, CStr(Procedure), S
        End If
    Next
End Sub


