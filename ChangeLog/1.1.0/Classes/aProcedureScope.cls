VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize  As aProcedureScope  :
' Fun ToString    As String           :
' Get Suggested   As String           :
' Sub MakePrivate                     :
' Sub MakePublic                      :
' Sub MakeSuggested                   :
'--------------------------------------------------
'@EndFolder ListOfProcedures


Option Explicit

Private oProcedure As New aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureScope
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Function ToString() As String
    Dim sTitle      As String
    sTitle = oProcedure.code.DeclarationClean
    If UCase(sTitle) Like UCase("*" & "Public " & "*" & oProcedure.Name & "*") Then
        ToString = "Public"
    ElseIf UCase(sTitle) Like UCase("*" & "Private " & "*" & oProcedure.Name & "*") Then
        ToString = "Private"
    Else
        Select Case oProcedure.Parent.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ToString = "Public"
            Case Else
                ToString = "Private"
        End Select
    End If
End Function

Public Property Get Suggested() As String
    Dim result      As String
    Dim cmCount As Long
    cmCount = oProcedure.Dependencies.CallerModules.count
    Select Case cmCount
        Case 0
            Suggested = "Unspecified"
            result = result & vbNewLine & "Procedure " & oProcedure.Name & " has unclarified scope in " & oProcedure.oWorkbook.Name
            result = result & vbNewLine & "It is called 0 times, so it may be Unused, Unassigned, or Called from elsewhere"
            Exit Property
        Case 1
            Suggested = "Private"
            result = result & vbNewLine & "Suggested scope for procedure " & oProcedure.Name & " is Private"
        Case Is > 1
            Suggested = "Public"
            result = result & vbNewLine & "Suggested scope for procedure " & oProcedure.Name & " is Public"
    End Select

    result = result & vbNewLine & "because it is used in " & cmCount & " module(s) of " & oProcedure.oWorkbook.Name
    Debug.Print result
End Property

Public Sub MakeSuggested()
    Dim idx         As Long: idx = oProcedure.lines.FirstOfDeclaration
    Dim s           As String: s = oProcedure.oModule.CodeModule.lines(idx, 1)
    Dim Suggestion  As String: Suggestion = Suggested
    If Suggestion = "Private" Or Suggestion = "Public" Then
        Select Case True
            Case s Like "*Public*" & oProcedure.Name & "*"
                MakePublic
            Case s Like "*Private*" & oProcedure.Name & "*"
                MakePrivate
            Case Else
                Dim sKind As String: sKind = oProcedure.KindAsString
                oProcedure.oModule.CodeModule.ReplaceLine idx, Trim(VBA.Replace(s, sKind, " " & Suggestion & " " & sKind, , 1, vbTextCompare))
        End Select
    End If
End Sub

Public Sub MakePrivate()
    Dim idx         As Long: idx = oProcedure.lines.FirstOfDeclaration
    Dim s           As String: s = oProcedure.code.DeclarationClean
    Dim Suggestion  As String: Suggestion = "Private"
    Select Case True
        Case s Like "*Public*" & oProcedure.Name & "*"
            Debug.Print oProcedure.Name
            oProcedure.oModule.CodeModule.ReplaceLine idx, VBA.Replace(s, "Public", "Private", , 1, vbTextCompare)
        Case s Like "*Private*" & oProcedure.Name & "*"

        Case Else
            Dim sKind As String: sKind = oProcedure.KindAsString
            Select Case sKind
                Case sKind Like "*Property*"
                    Debug.Print oProcedure.Name
                    oProcedure.oModule.CodeModule.ReplaceLine idx, VBA.Replace(s, "Property", Suggestion & " " & "Property", , 1, vbTextCompare)
                Case Else
                    Debug.Print oProcedure.Name
                    oProcedure.oModule.CodeModule.ReplaceLine idx, VBA.Replace(s, sKind, Suggestion & " " & sKind, , 1, vbTextCompare)
            End Select
    End Select
End Sub

Public Sub MakePublic()
    Dim idx         As Long: idx = oProcedure.lines.FirstOfDeclaration
    Dim s           As String: s = oProcedure.code.DeclarationClean
    Dim Suggestion  As String: Suggestion = "Public"
    Select Case True
        Case s Like "*Public*" & oProcedure.Name & "*"

        Case s Like "*Private*" & oProcedure.Name & "*"
            Debug.Print oProcedure.Name
            oProcedure.oModule.CodeModule.ReplaceLine idx, VBA.Replace(s, "Private", "Public", , 1, vbTextCompare)
        Case Else
            Dim sKind As String: sKind = oProcedure.KindAsString
            Debug.Print oProcedure.Name
            oProcedure.oModule.CodeModule.ReplaceLine idx, Trim(VBA.Replace(s, sKind, " " & Suggestion & " " & sKind, , 1, vbTextCompare))
    End Select
End Sub

