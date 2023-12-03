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
    Set Format = New aProcedureFormat
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
    Dim s           As String: s = code.DeclarationClean
    Select Case True
        Case InStr(1, s, "Get " & ProcedureName) > 0: KindAsString = "Get"
        Case InStr(1, s, "Let " & ProcedureName) > 0: KindAsString = "Let"
        Case InStr(1, s, "Set " & ProcedureName) > 0: KindAsString = "Set"
        Case Else
            If InStr(1, s, "Function " & ProcedureName) > 0 Then
                KindAsString = "Function"
            ElseIf InStr(1, s, "Sub " & ProcedureName) > 0 Then
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

Public Property Get returnTypeUnassigned() As Boolean
    returnTypeUnassigned = (returnType = "Variant" And Not code.Contains("set " & Name, True, False, False))
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



Function testsub()

End Function

Public Sub CreateTest(Optional forObject As Boolean)
    Dim arg         As String
    arg = arguments.MultiLine
    Dim Procedures  As Collection
    Set Procedures = aProject.Initialize(oWorkbook).ProceduresNames(True, False, False, False)

    Dim s           As String: s = vbNewLine
    If Not aCollection.Init(Procedures).Contains(, "test" & ProcedureName) Then
        
        If KindAsString = "Sub" Then
            s = Join( _
                    Array( _
                        s, _
                        "Function test" & ProcedureName & "()", _
                        IIf(arg <> "", "call ", "") & arg, _
                        "dim x,y,z", _
                        "test" & ProcedureName & "= (x=1 and y=2) or (z=3)", _
                        "End Function"), _
                    vbNewLine)
        Else
            s = Join( _
                    Array(s, _
                        "Function test" & ProcedureName & "()", _
                        "Dim Something As " & IIf(forObject, "New ", "") & returnType, _
                        IIf(forObject, "Set ", "") & "Something = " & arg, _
                        "dim x,y,z", _
                        "test" & ProcedureName & "= (x=1 and y=2) or (z=3)", _
                        "End Function"), _
                    vbNewLine)
        End If
        oModule.CodeModule.InsertLines lines.first, s
        Debug.Print "Created test" & ProcedureName & " in " & oModule.Name
    Else

    End If
End Sub

Public Sub Delete()
    Dim targetFile As String
    targetFile = Environ("TEMP") & "\" & ProcedureName & ".txt"
    TxtOverwrite targetFile, code.All
    Debug.Print "made a backup of " & ProcedureName
    Debug.Print "followlink " & """" & targetFile & """"
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




