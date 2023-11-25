VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aModuleEnumItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize    As aModuleEnumItem  :
' Get Body          As String           :
' Get Name          As String           :
' Get countoflines  As Long             :
' Get firstline     As Long             :
' Get index         As String           :
' Get lastline      As Long             :
' Let Name                              :
' Let index                             :
' Sub AssignValues                      :
' Sub ToCase                            :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private Type TaModuleEnumItem
    Name As String
    index As Long
End Type

Private this As TaModuleEnumItem

Private oModule As aModule

Public Function Initialize(Module As aModule) As aModuleEnumItem
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

Public Property Get countoflines() As Long
    countoflines = lastline - firstline + 1
End Property

Public Property Get Body() As String
    Body = oModule.Component.CodeModule.lines(firstline + 1, countoflines - 2)
End Property

Private Sub replaceBody(S As String)
    With oModule.Component.CodeModule
        .DeleteLines firstline + 1, countoflines - 2
        .InsertLines lastline, S
    End With
End Sub

Public Sub AssignValues(Optional ToThePower As Boolean = True)
    Dim arr: arr = cleanArray(Split(Body, vbNewLine))
    Dim i           As Long
    For i = 0 To UBound(arr)
        arr(i) = Split(arr(i), "=")(0)
        arr(i) = Space(4) & arr(i)
        If ToThePower Then
            arr(i) = arr(i) & "= 2 ^ " & i
        Else
            arr(i) = arr(i) & "= " & i + 1
        End If
    Next
    replaceBody Join(arr, vbNewLine)
End Sub

Public Sub ToCase()
    Dim arr: arr = cleanArray(Split(Body, vbNewLine))
    Dim out         As String: out = "Select case Variable " & vbNewLine
    Dim i           As Long
    For i = 0 To UBound(arr)
        arr(i) = Split(arr(i), "=")(0)
        out = out & Space(4) & "Case is = " & Name & "." & arr(i) & vbNewLine
    Next
    
    out = out & vbNewLine & "End Select"
    Debug.Print "The following was copied to clipboard:"
    Debug.Print String(20, "-")
    Debug.Print out
    CLIP out
End Sub


