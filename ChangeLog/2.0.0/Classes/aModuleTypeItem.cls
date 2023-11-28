VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aModuleTypeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize    As aModuleTypeItem  :
' Get Body          As String           :
' Get Name          As String           :
' Get countoflines  As Long             :
' Get firstline     As Long             :
' Get index         As String           :
' Get lastline      As Long             :
' Let Name                              :
' Let index                             :
' Sub AssignValues                      :
'--------------------------------------------------
'@EndFolder ListOfProcedures
Option Explicit


Private Type TaModuleEnumItem
    Name As String
    index As Long
End Type

Private this As TaModuleEnumItem

Private oModule As aModule

Public Function Initialize(Module As aModule) As aModuleTypeItem
    Set oModule = Module
    Set Initialize = Me
End Function

Public Property Let Name(val As String)
    this.Name = val
End Property
Public Property Get Name() As String
    Name = this.Name
End Property
Public Property Let index(val As String)
    this.index = val
End Property
Public Property Get index() As String
    index = this.index
End Property

Public Property Get firstline() As Long
    firstline = oModule.LineLike("Enum " & this.Name, True, True)
End Property

Public Property Get lastline() As Long
    Dim i As Long
    With oModule.Component.CodeModule
        For i = firstline To .CountOfDeclarationLines
            If Trim(.lines(i, 1)) = "End Enum" Then
                lastline = i
                Exit For
            End If
        Next
    End With
End Property

Public Property Get countOfLines() As Long
    countOfLines = lastline - firstline + 1
End Property

Public Property Get Body() As String
    Body = oModule.Component.CodeModule.lines(firstline + 1, countOfLines - 2)
End Property

Private Sub replaceBody(S As String)
    With oModule.Component.CodeModule
        .DeleteLines firstline + 1, countOfLines - 2
        .InsertLines lastline, S
    End With
End Sub

Public Sub AssignValues(Optional ToThePower As Boolean = True)
    Dim arr: arr = cleanArray(Split(Body, vbNewLine))
    Dim i           As Long
    For i = 0 To UBound(arr)

    Next
    replaceBody Join(arr, vbNewLine)
End Sub


