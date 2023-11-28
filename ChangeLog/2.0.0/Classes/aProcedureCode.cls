VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize         As aProcedureCode    :
' Fun Inject             As aProcedureInject  :
' Fun lines              As aProcedureLines   :
' Get All                As Variant           :
' Get Body               As Variant           :
' Get BodyAfterComments  As Variant           :
' Get Contains           As Boolean           :
' Get ContainsInBody     As String            :
' Get ContainsInHeader   As String            :
' Get Declaration        As Variant           :
' Get DeclarationClean   As Variant           :
' Get Header             As Variant           :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureCode
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Function lines() As aProcedureLines
    Set lines = oProcedure.lines
End Function

Public Function Inject() As aProcedureInject
    Set Inject = oProcedure.Inject
End Function

Public Property Get All()
    All = oProcedure.Parent.CodeModule.lines(lines.first, lines.count)
End Property
Public Property Get Body()
    Body = oProcedure.Parent.CodeModule.lines(lines.FirstOfBody, lines.CountOfBody)
End Property
Public Property Get BodyAfterComments()
    BodyAfterComments = oProcedure.Parent.CodeModule.lines(lines.FirstOfBodyAfterComments, lines.CountOfBody)
End Property

Public Property Get Header()
    Header = oProcedure.Parent.CodeModule.lines(lines.FirstOfHeader, lines.CountOfHeaderLines)
End Property
Public Property Get Declaration()
    Declaration = oProcedure.Parent.CodeModule.lines(lines.FirstOfDeclaration, lines.CountOfDeclarationLines)
End Property
Public Property Get DeclarationClean()
    If InStr(1, Declaration, "_") = 0 Then
        DeclarationClean = Trim(Declaration)
    Else
        DeclarationClean = VBA.Replace(Join(ArrayRemoveEmptyElements(ArrayTrim(Split(Declaration, "_" & vbNewLine))), " "), "( ", "(")
    End If
End Property

Public Property Get Contains( _
                        this As String, _
                        WholeWord As Boolean, _
                        MatchCase As Boolean, _
                        PatternSearch As Boolean) As Boolean
                        
    Contains = oProcedure.Parent.CodeModule.Find( _
                                                this, _
                                                1, _
                                                lines.first, _
                                                lines.last, _
                                                10000, _
                                                WholeWord, _
                                                MatchCase, _
                                                PatternSearch)
End Property

Public Property Get ContainsInHeader( _
                                this As String, _
                                WholeWord As Boolean, _
                                MatchCase As Boolean, _
                                PatternSearch As Boolean) As String

    ContainsInHeader = oProcedure.Parent.CodeModule.Find( _
                                                this, _
                                                1, _
                                                lines.FirstOfHeader, _
                                                lines.LastOfHeader, _
                                                10000, _
                                                WholeWord, _
                                                MatchCase, _
                                                PatternSearch)
End Property

Public Property Get ContainsInBody( _
                                this As String, _
                                WholeWord As Boolean, _
                                MatchCase As Boolean, _
                                PatternSearch As Boolean) As String

    ContainsInBody = oProcedure.Parent.CodeModule.Find( _
                                            this, _
                                            1, _
                                            lines.FirstOfBody, _
                                            lines.LastOfBody, _
                                            10000, _
                                            WholeWord, _
                                            MatchCase, _
                                            PatternSearch)
End Property

