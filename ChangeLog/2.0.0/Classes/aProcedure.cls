VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Active            As aProcedure                  :
' Fun Code              As aProcedureCode              :
' Fun CustomProperties  As aProcedureCustomProperties  :
' Fun Dependencies      As aProcedureDependencies      :
' Fun Folder            As aProcedureFolder            :
' Fun Format            As aProcedureFormat            :
' Fun Initialize        As aProcedure                  :
' Fun Inject            As aProcedureInject            :
' Fun Move              As aProcedureMove              :
' Fun Scope             As aProcedureScope             :
' Fun Variables         As aProcedureVariables         :
' Fun arguments         As aProcedureArguments         :
' Fun lines             As aProcedureLines             :
' Get KindAsLong        As Long                        :
' Get KindAsString      As String                      :
' Get Name              As String                      :
' Get Parent            As VBComponent                 :
' Get returnType        As String                      :
' Sub Activate                                         :
' Sub CreateCaller                                     :
' Sub CreateTest                                       :
' Sub Delete                                           :
' Sub Replace                                          :
'--------------------------------------------------
'@EndFolder ListOfProcedures



'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aProcedure
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:18    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private ProcedureName    As String
Public oWorkbook         As Workbook
Public oModule           As VBComponent
Public oProcKind         As vbext_ProcKind
'
'Private Type tAddons
'    Code              As aProcedureCode
'    Move              As aProcedureMove
'    Format            As aProcedureFormat
'    Lines             As aProcedureLines
'    Inject            As aProcedureInject
'    Dependencies      As aProcedureDependencies
'    arguments         As aProcedureArguments
'    Variables         As aProcedureVariables
'    CustomProperties  As aProcedureCustomProperties
'    Scope             As aProcedureScope
'    Folder            As aProcedureFolder
'End Type
'
'Private Addons As tAddons

Public Function Initialize( _
                    Optional TargetWorkbook As Workbook, _
                    Optional Module As VBComponent, _
                    Optional Procedure As String, _
                    Optional procKind As vbext_ProcKind) As aProcedure
                    
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    ProcedureName = Procedure
    Set oWorkbook = TargetWorkbook
    Set oModule = Module
    oProcKind = procKind
    Set Initialize = Me
End Function

Public Function Active() As aProcedure
    Dim Module      As VBComponent: Set Module = ActiveModule
    Dim Procedure   As String
    Dim i           As Long
    Dim procKind    As vbext_ProcKind
    Dim lineNum     As Long: lineNum = aCodeModule.Initialize(Module).firstRow
    For i = 0 To 4
        procKind = i
        Procedure = Module.CodeModule.ProcOfLine(lineNum, procKind)
        If Procedure <> vbNullString Then
            Initialize ActiveCodepaneWorkbook, Module, Procedure, procKind
            Set Active = Me
            Exit For
        End If
    Next
End Function


Public Function code() As aProcedureCode
    Set code = New aProcedureCode
    code.Initialize Me
End Function
Public Function Move() As aProcedureMove
    Set Move = New aProcedureMove
    Move.Initialize Me
End Function
Public Function Format() As aProcedureFormat
    Format = New aProcedureFormat
    Format.Initialize Me
End Function
Public Function lines() As aProcedureLines
    Set lines = New aProcedureLines
    lines.Initialize Me
End Function
Public Function Inject() As aProcedureInject
    Set Inject = New aProcedureInject
    Inject.Initialize Me
End Function
Public Function Dependencies() As aProcedureDependencies
    Set Dependencies = New aProcedureDependencies
    Dependencies.Initialize Me
End Function
Public Function arguments() As aProcedureArguments
    Set arguments = New aProcedureArguments
    arguments.Initialize Me
End Function
Public Function Variables() As aProcedureVariables
    Set Variables = New aProcedureVariables
    Variables.Initialize Me
End Function
Public Function CustomProperties() As aProcedureCustomProperties
    Set CustomProperties = New aProcedureCustomProperties
    CustomProperties.Initialize Me
End Function
Public Function Scope() As aProcedureScope
    Set Scope = New aProcedureScope
    Scope.Initialize Me
End Function
Public Function Folder() As aProcedureFolder
    Set Folder = New aProcedureFolder
    Folder.Initialize Me
End Function

'Public Function Code() As aProcedureCode
'    If Addons.Code Is Nothing Then
'        Set Addons.Code = New aProcedureCode
'        Addons.Code.Initialize Me
'    End If
'    Set Code = Addons.Code
'End Function
'Public Function Move() As aProcedureMove
'    If Addons.Move Is Nothing Then
'        Set Addons.Move = New aProcedureMove
'        Addons.Move.Initialize Me
'    End If
'    Set Move = Addons.Move
'End Function
'Public Function Format() As aProcedureFormat
'    If Addons.Format Is Nothing Then
'        Set Addons.Format = New aProcedureFormat
'        Addons.Format.Initialize Me
'    End If
'    Set Format = Addons.Format
'End Function
'Public Function Lines() As aProcedureLines
'    If Addons.Lines Is Nothing Then
'        Set Addons.Lines = New aProcedureLines
'        Addons.Lines.Initialize Me
'    End If
'    Set Lines = Addons.Lines
'End Function
'Public Function Inject() As aProcedureInject
'    If Addons.Inject Is Nothing Then
'        Set Addons.Inject = New aProcedureInject
'        Addons.Inject.Initialize Me
'    End If
'    Set Inject = Addons.Inject
'End Function
'Public Function Dependencies() As aProcedureDependencies
'    If Addons.Dependencies Is Nothing Then
'        Set Addons.Dependencies = New aProcedureDependencies
'        Addons.Dependencies.Initialize Me
'    End If
'    Set Dependencies = Addons.Dependencies
'End Function
'Public Function arguments() As aProcedureArguments
'    If Addons.arguments Is Nothing Then
'        Set Addons.arguments = New aProcedureArguments
'        Addons.arguments.Initialize Me
'    End If
'    Set arguments = Addons.arguments
'End Function
'Public Function Variables() As aProcedureVariables
'    If Addons.Variables Is Nothing Then
'        Set Addons.Variables = New aProcedureVariables
'        Addons.Variables.Initialize Me
'    End If
'    Set Variables = Addons.Variables
'End Function
'Public Function CustomProperties() As aProcedureCustomProperties
'    If Addons.CustomProperties Is Nothing Then
'        Set Addons.CustomProperties = New aProcedureCustomProperties
'        Addons.CustomProperties.Initialize Me
'    End If
'    Set CustomProperties = Addons.CustomProperties
'End Function
'Public Function Scope() As aProcedureScope
'    If Addons.Scope Is Nothing Then
'        Set Addons.Scope = New aProcedureScope
'        Addons.Scope.Initialize Me
'    End If
'    Set Scope = Addons.Scope
'End Function
'Public Function Folder() As aProcedureFolder
'    If Addons.Folder Is Nothing Then
'        Set Addons.Folder = New aProcedureFolder
'        Addons.Folder.Initialize Me
'    End If
'    Set Folder = Addons.Folder
'End Function


Public Sub Activate()
    oModule.Activate
    Dim StartPos    As Long: StartPos = InStr(1, oModule.CodeModule.lines(lines.FirstOfDeclaration, 1), Me.Name, vbTextCompare)
    oModule.CodeModule.CodePane.SetSelection lines.FirstOfDeclaration, _
                                             StartPos, _
                                             lines.FirstOfDeclaration, _
                                             StartPos + Len(Me.Name)
End Sub

Public Property Get Name() As String
    Name = ProcedureName
End Property

Public Property Get Parent() As VBComponent
    Set Parent = oModule
End Property

Public Property Get KindAsLong() As Long
    KindAsLong = oProcKind
End Property

Public Property Get KindAsString() As String
    Dim S           As String: S = code.DeclarationClean
    Select Case True
        Case InStr(1, S, "Get " & ProcedureName) > 0: KindAsString = "Get"
        Case InStr(1, S, "Let " & ProcedureName) > 0: KindAsString = "Let"
        Case InStr(1, S, "Set " & ProcedureName) > 0: KindAsString = "Set"
        Case Else
            If InStr(1, S, "Function " & ProcedureName) > 0 Then
                KindAsString = "Function"
            ElseIf InStr(1, S, "Sub " & ProcedureName) > 0 Then
                KindAsString = "Sub"
            End If
    End Select
End Property

Public Property Get returnType() As String
    Select Case KindAsString
    Case "Sub", "Let", "Set": Exit Property
    End Select
    Dim title       As String: title = code.DeclarationClean
    If title Like "*) As *" Then
        returnType = Mid(title, InStrRev(title, ") As ") + Len(") As "))
    Else
        returnType = "Variant"
    End If
End Property

Sub CreateCaller(rng As Range)
    Dim shp         As Shape
    Set shp = ActiveSheet.Shapes.AddShape _
            (msoShapeRoundedRectangle, 1, 1, 500, 10)
    With shp.ThreeD
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 6
        .BevelTopDepth = 6
    End With
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With

    With shp
        .OnAction = "'" & oWorkbook.Name & "'!" & ProcedureName
        .Name = "Run_" & ProcedureName
        .TextFrame2.TextRange.text = ProcedureName
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame2.WordWrap = msoFalse
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .Left = Selection.Left
        .Top = Selection.Top
    End With
End Sub

Public Sub CreateTest()
    Dim arg         As String
    arg = arguments.MultiLine
    Dim Procedures  As Collection
    Set Procedures = aProject.Initialize(oWorkbook).Procedures(True, False, False, False)

    Dim S           As String: S = vbNewLine
    If Not aCollection.Init(Procedures).Contains(, "test" & ProcedureName) Then
        If KindAsString = "Sub" Then
            S = S & "Sub test" & ProcedureName & "()" & vbNewLine
            S = S & "call " & arg & vbNewLine
            S = S & "End Sub"
        Else
            S = S & "Function test" & ProcedureName & "()" & vbNewLine
            S = S & "Something = " & arg & vbNewLine
            S = S & "End Function"
        End If
        oModule.CodeModule.InsertLines lines.first, S
        Debug.Print "Created test" & ProcedureName & " in " & oModule.Name
    Else

    End If
End Sub

Public Sub Delete()
    Dim targetFile As String
    targetFile = Environ("TEMP") & "\" & ProcedureName & ".txt"
    TxtOverwrite targetFile, code.All
    Debug.Print "made a backup of " & ProcedureName
    Debug.Print "followlink " & targetFile
    oModule.CodeModule.DeleteLines lines.first, lines.count
End Sub

Public Sub Replace(newCode As String)
    Dim targetLine   As Long
    targetLine = lines.first
    Delete
    oModule.CodeModule.InsertLines targetLine, newCode
End Sub
Private Sub Class_Terminate()
    ReSetVariables
End Sub

Private Sub ReSetVariables()
    Set oWorkbook = Nothing
    Set oModule = Nothing
    ProcedureName = ""
End Sub




