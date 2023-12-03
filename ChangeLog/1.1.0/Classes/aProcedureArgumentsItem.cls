VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureArgumentsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize       As aProcedureArgumentsItem  :
' Get DefaultValue     As Variant                  :
' Get IsByRef          As Boolean                  :
' Get IsByVal          As Boolean                  :
' Get IsOptional       As Boolean                  :
' Get IsParamArray     As Boolean                  :
' Get IsType           As String                   :
' Get Name             As String                   :
' Get hasDefaultValue  As Boolean                  :
' Let DefaultValue                                 :
' Let IsByRef                                      :
' Let IsByVal                                      :
' Let IsOptional                                   :
' Let IsParamArray                                 :
' Let IsType                                       :
' Let Name                                         :
' Let hasDefaultValue                              :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Private Type TaProcedureArgumentsItem
     index As Long
     Name As String
     IsByRef As Boolean
     IsByVal As Boolean
     IsOptional As Boolean
     hasDefaultValue As Boolean
     DefaultValue As Variant
     IsParamArray As Boolean
     IsType As String
     OriginalString As String
End Type

Private this As TaProcedureArgumentsItem

Public Function Initialize(Procedure As aProcedure) As aProcedureArgumentsItem
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Property Let index(val As String)
    this.index = val
End Property
Public Property Let Name(val As String)
    this.Name = val
End Property
Public Property Let IsByRef(val As Boolean)
    this.IsByRef = val
End Property
Public Property Let IsByVal(val As Boolean)
    this.IsByVal = val
End Property
Public Property Let IsOptional(val As Boolean)
    this.IsOptional = val
End Property
Public Property Let DefaultValue(val As Variant)
    this.DefaultValue = val
End Property
Public Property Let hasDefaultValue(val As Boolean)
    this.hasDefaultValue = val
End Property
Public Property Let IsParamArray(val As Boolean)
    this.IsParamArray = val
End Property
Public Property Let IsType(val As String)
    this.IsType = val
End Property



Public Property Let OriginalString(val As String)
    this.OriginalString = val
End Property

Public Sub SetType(sTypeName As String)
    Dim parenthesesPattern As String
    If InStr(1, this.OriginalString, "()", vbTextCompare) > 0 Then parenthesesPattern = "\(\)"
    Dim asPattern As String
    If InStr(1, this.OriginalString, " As ", vbTextCompare) > 0 Then asPattern = " As\s+\w+"
    Dim regexPattern As String
    regexPattern = "\b" & this.Name & parenthesesPattern & asPattern
    oProcedure.Replace RegExpReplace(oProcedure.code.All, regexPattern, this.Name & IIf(parenthesesPattern <> "", "()", "") & " As " & sTypeName, 1)
    oProcedure.Parent.CodeModule.CodePane.SetSelection oProcedure.lines.FirstOfDeclaration, 1, oProcedure.lines.FirstOfDeclaration, 1
End Sub

Public Property Get OriginalString() As String
    OriginalString = this.OriginalString
End Property



Public Property Get index() As String
    index = this.index
End Property
Public Property Get Name() As String
    Name = this.Name
End Property
Public Property Get IsByRef() As Boolean
    IsByRef = this.IsByRef
End Property
Public Property Get IsByVal() As Boolean
    IsByVal = this.IsByVal
End Property
Public Property Get IsOptional() As Boolean
    IsOptional = this.IsOptional
End Property
Public Property Get hasDefaultValue() As Boolean
    hasDefaultValue = this.hasDefaultValue
End Property
Public Property Get DefaultValue() As Variant
    DefaultValue = this.DefaultValue
End Property
Public Property Get IsParamArray() As Boolean
    IsParamArray = this.IsParamArray
End Property
Public Property Get IsType() As String
    IsType = this.IsType
End Property

