VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureVariablesItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize       As aProcedureVariablesItem  :
' Get IsType           As String                   :
' Get Line             As Long                     :
' Get Name             As String                   :
' Get isAssignedValue  As Boolean                  :
' Let IsType                                       :
' Let Line                                         :
' Let Name                                         :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Private Type TaProcedureVariablesItem
    Line As Long
    Name As String
    IsType As String
End Type

Private this As TaProcedureVariablesItem

Public Function Initialize(Procedure As aProcedure) As aProcedureVariablesItem
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Property Get isAssignedValue() As Boolean
    isAssignedValue = oProcedure.code.Contains("*" & this.Name & "* = *", False, False, True) _
                      Or oProcedure.code.Contains("*" & this.Name & "* As New *", False, False, True)
End Property

Public Property Get Line() As Long
    Line = this.Line
End Property

Public Property Let Line(ByVal val As Long)
    this.Line = val
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal val As String)
    this.Name = val
End Property

Public Property Get IsType() As String
    IsType = this.IsType
End Property

Public Property Let IsType(ByVal val As String)
    this.IsType = val
End Property

