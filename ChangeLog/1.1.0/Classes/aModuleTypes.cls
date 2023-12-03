VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aModuleTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Active      As aModuleTypeItem  :
' Fun Initialize  As aModuleTypes     :
' Get Items       As Variant          :
'--------------------------------------------------
'@EndFolder ListOfProcedures
Option Explicit

Private oModule As aModule

Public Function Initialize(Module As aModule) As aModuleTypes
    Set oModule = Module
    Set Initialize = Me
End Function

Public Function Active() As aModuleTypeItem
    Dim acm As aCodeModule
    Set acm = aCodeModule.Active
    If acm.firstRow > oModule.Component.CodeModule.CountOfDeclarationLines Then
        Toast "Cursor not inside an Type"
        End
    End If
    
    Dim Line        As String
    Dim i           As Long
    For i = aCodeModule.Initialize(oModule.Component).firstRow To 1 Step -1
        Line = Trim(oModule.Component.CodeModule.lines(i, 1))
        If Line Like "Type *" Then
            Set Active = Items(Trim(Split(Line, "Type ")(1)))
            Exit Function
        End If
    Next
End Function

Public Property Get Items(Optional index As Variant)
    Dim coll As Collection
    Set coll = myItems
    Dim this As aModuleEnumItem '<-----
    If IsMissing(index) Then
        Set Items = coll
    ElseIf IsNumeric(index) Then
        Set Items = coll(index)
    ElseIf TypeName(index) = "String" Then
        For Each this In coll
            If this.Name = index Then
                Set Items = this
                Exit For
            End If
        Next
    End If
End Property

Private Function myItems() As Collection
    Dim coll As Collection
    Set coll = New Collection
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    Dim matches As Object
    Dim match As Object
    
    regex.Global = True
    regex.IgnoreCase = True
    regex.MultiLine = True
    regex.pattern = "Type\s+(\w+)\s*[\s\S]*?End Type"
    
    Set matches = regex.Execute(oModule.Header)
    
    Dim this As aModuleEnumItem '<-----
    Dim i As Long
    For Each match In matches
        i = i + 1
        Set this = New aModuleEnumItem
        this.Initialize oModule
        this.Name = Split(Trim(Split(match, vbNewLine)(0)), " ")(1)
        this.index = i
        coll.Add this
    Next match
    Set myItems = coll
End Function

