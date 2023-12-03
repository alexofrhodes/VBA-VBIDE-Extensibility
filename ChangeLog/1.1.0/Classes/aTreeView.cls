VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun GetLevel            As Integer      :
' Fun Init                As aTreeView    :
' Fun ToArray             As Variant      :
' Fun TreeviewArrayPaths  As Variant '1d  :
' Fun columnCount         As Variant      :
' Fun rowCount            As Variant      :
' Sub ActivateProjectElement              :
' Sub ApplyStandardStyle                  :
' Sub ChildrenCheck                       :
' Sub CollapseAll                         :
' Sub ExpandAll                           :
' Sub FilterTV                            :
' Sub FindCodeEverywhere                  :
' Sub ImageListLoadProjectIcons           :
' Sub LoadRange                           :
' Sub LoadTreeArray                       :
' Sub LoadVBProjects                      :
' Sub RemoveEmpty                         :
' Sub SelectNextNode                      :
' Sub SelectNodes                         :
' Sub SelectPreviousNode                  :
' Sub TreeviewArrayAppendPaths            :
' Sub TreeviewAssignProjectImages         :
' Sub clear                               :
'--------------------------------------------------
'@EndFolder ListOfProcedures

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aTreeView
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:19    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private TV          As TreeView

Private Enum tvImages
    tvProject = 1
    tvSheet = 2
    tvForm = 3
    tvModule = 4
    tvClass = 5
    tvMacro = 6
    tvText = 7
End Enum
Const iTvProject = 1

Public Function Init(TargetTreeview) As aTreeView
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Set TV = TargetTreeview
    Set Init = Me
End Function

'Sub TVTestVarious()
'    zTest.Show
'    Dim nod As Node
'    Dim tv As TreeView
'    Set tv = zTest.TreeView1
'    TreeviewPopulateFromRange tv, ActiveSheet.Range("A1"), True, False
'    For Each nod In tv.Nodes
'        dp nod.TEXT & vbTab & TreeviewGetLevel(nod)
'    Next
'    dp TreeviewArray(tv)
'    Stop
'    dp TreeviewArrayPaths(tv)
'    Stop
'    TreeviewFilter tv, "*\*\*\*", True
'    dp TreeviewArrayPaths(tv)
'    TreeviewFoldingExpandAll tv
'End Sub

Public Sub RemoveEmpty(ColumnLevel As Long)
    '@AssignedModule aTreeView
    '@INCLUDE PROCEDURE dp
    '@INCLUDE CLASS aTreeView
    Dim i           As Long
    Dim nod         As Node
    For i = TV.Nodes.count To 1 Step -1
        Set nod = TV.Nodes(i)
        If GetLevel(nod) = ColumnLevel Then
            If nod.Children = 0 Then
                TV.Nodes.Remove nod.index
            End If
        End If
    Next
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 20-12-2022 07:43    Alex                (TreeviewFilter)

Public Sub FilterTV(txtLike As String, MatchCase As Boolean)
    'example
    'TreeviewFilter tv, "*\*\*\*", True
    'each \ indicates a sublevel
    '@AssignedModule aTreeView
    '@INCLUDE PROCEDURE ArrayFilterLike
    '@INCLUDE CLASS aTreeView

    Dim var
    var = ArrayFilterLike(TreeviewArrayPaths(TV), txtLike, MatchCase)
    clear
    Dim element
    Dim nod         As Node
    On Error Resume Next
    Dim i           As Long
    For Each element In var
        For i = 0 To UBound(Split(element, "\"))
            If i = 0 Then
                Set nod = TV.Nodes.Add(, , , Split(element, "\")(i))
            Else
                Set nod = TV.Nodes.Add(nod, tvwChild, , Split(element, "\")(i))
            End If
        Next
    Next
    On Error GoTo 0
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 19-12-2022 15:09    Alex                (TVColumnsCount)

Public Function columnCount()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Dim Output      As Long
    Dim counter     As Long
    Dim nod         As Node
    For Each nod In TV.Nodes
        counter = GetLevel(nod)
        If counter > Output Then Output = Output + 1
    Next
    columnCount = Output
End Function

Public Function rowCount()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    rowCount = TV.Nodes.count
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 19-07-2023 10:23    Alex                added base 1 : ReDim output(1 To rowCount, 1 To ColumnCount)

Function ToArray() As Variant
    '@LastModified 2307191023
    'Modified from https://www.excelforum.com/excel-programming-vba-macros/1102518-transfer-all-nodes-from-a-userform-treeview-into-a-treeview-on-a-worksheet-with-vba.html
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Dim nROOT       As Node
    Dim nX          As Node
    Dim irow        As Integer

    On Error GoTo catch

    If TV.Nodes.count = 0 Then
        MsgBox "No nodes to export...", vbExclamation, "Nothing to do"
        Stop
    End If

    ' Start at first node
    Set nX = TV.Nodes(1)
    ' Get the first root node
    Set nROOT = nX.Root.FirstSibling
    'Create array placeholder
    Dim Output()
    ReDim Output(1 To rowCount, 1 To columnCount)
    'Drill down and export
    Do
        TreeviewArrayAppend 0, irow, nROOT, Output, True
        Set nROOT = nROOT.Next
    Loop While Not (nROOT Is Nothing)

    ToArray = Output
Finally:
    'Cleanup

    'Exit
    On Error GoTo 0
    Exit Function
catch:
    MsgBox "ERROR: " & Err.Description, vbExclamation, "Error"
    Resume Finally
End Function

Private Sub TreeviewArrayAppend(ByVal lngLevel As Integer, _
        ByRef lngRow As Integer, _
        ByRef N As Node, _
        ByRef TargetArray As Variant, _
        ByVal includeChildren As Boolean)
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView

    TargetArray(lngRow + 1, lngLevel + 1) = N.text
    lngRow = lngRow + 1
    ' If there are child nodes then call recusively
    lngLevel = lngLevel + 1
    If includeChildren Then
        If N.Children > 0 Then
            Dim nC  As Node
            Set nC = N.child
            Do
                TreeviewArrayAppend lngLevel, lngRow, nC, TargetArray, True
                Set nC = nC.Next
            Loop While Not (nC Is Nothing)
        End If
    End If
End Sub
'--------------------------------------------

Public Function TreeviewArrayPaths() As Variant  '1d
    'Modified from https://www.excelforum.com/excel-programming-vba-macros/1102518-transfer-all-nodes-from-a-userform-treeview-into-a-treeview-on-a-worksheet-with-vba.html
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Dim nROOT       As Node
    Dim nX          As Node
    Dim irow        As Integer

    On Error GoTo catch

    If TV.Nodes.count = 0 Then
        MsgBox "No nodes to export...", vbExclamation, "Nothing to do"
        Stop
    End If

    ' Start at first node
    Set nX = TV.Nodes(1)
    ' Get the first root node
    Set nROOT = nX.Root.FirstSibling
    'Create array placeholder
    Dim Output()
    ReDim Output(rowCount)
    'Drill down and export
    Do
        TreeviewArrayAppendPaths 0, irow, nROOT, Output
        Set nROOT = nROOT.Next
    Loop While Not (nROOT Is Nothing)

    TreeviewArrayPaths = Output
Finally:
    'Cleanup

    'Exit
    On Error GoTo 0
    Exit Function
catch:
    MsgBox "ERROR: " & Err.Description, vbExclamation, "Error"
    Resume Finally
End Function

Public Sub TreeviewArrayAppendPaths(ByVal lngLevel As Integer, _
        ByRef lngRow As Integer, _
        ByRef N As Node, _
        ByRef TargetArray As Variant)
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView

    TargetArray(lngRow) = N.fullPath
    lngRow = lngRow + 1
    ' If there are child nodes then call recusively
    lngLevel = lngLevel + 1
    If N.Children > 0 Then
        Dim nC      As Node
        Set nC = N.child
        Do
            TreeviewArrayAppendPaths lngLevel, lngRow, nC, TargetArray
            Set nC = nC.Next
        Loop While Not (nC Is Nothing)
    End If
End Sub

Public Sub SelectNextNode()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    If Not TV.SelectedItem.child Is Nothing Then
        TV.SelectedItem.child.Selected = True
    ElseIf Not TV.SelectedItem.Next Is Nothing Then
        TV.SelectedItem.Next.Selected = True
    Else
        If Not TV.SelectedItem.Parent.Next Is Nothing Then
            TV.SelectedItem.Parent.Next.Selected = True
        Else
            TV.SelectedItem.Parent.Parent.Next.Selected = True
        End If
    End If
End Sub

Public Sub SelectPreviousNode()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    If Not TV.SelectedItem.Previous Is Nothing Then
        If Not TV.SelectedItem.Previous.child Is Nothing Then
            TV.SelectedItem.Previous.child.LastSibling.Selected = True
            Do While Not TV.SelectedItem.child Is Nothing
                TV.SelectedItem.child.LastSibling.Selected = True
            Loop
        Else
            TV.SelectedItem.Previous.Selected = True
        End If
    ElseIf Not TV.SelectedItem.Parent Is Nothing Then
        TV.SelectedItem.Parent.Selected = True
    End If
End Sub

Public Sub SelectNodes(SingleSelect As Boolean, lvl1crit As String, Optional CriteriaByLevel As Variant)
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Dim nd          As Node
    For Each nd In TV.Nodes
        If nd.text = lvl1crit Then
            nd.Selected = True
            nd.Expanded = True
            If SingleSelect = True Then Exit For
        End If
    Next
    Dim x           As Long
    x = nd.index + 1
    Dim crit
    Dim i           As Long
    For Each crit In CriteriaByLevel
        For i = x To TV.Nodes.count
            If TV.Nodes.Item(i).text = crit Then
                TV.Nodes.Item(i).Selected = True
                TV.Nodes.Item(i).Expanded = True
                If SingleSelect = True Then Exit For
            End If
        Next
        x = i + 1
    Next
End Sub

Public Sub ApplyStandardStyle()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    With TV
        .Sorted = False
        .Appearance = ccFlat
        .LineStyle = tvwRootLines
        .Style = tvwTreelinesPlusMinusPictureText
        .Font.Size = 10
        .indentation = 2
    End With
End Sub

Sub ExpandAll()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Dim i           As Long
    For i = 1 To TV.Nodes.count
        TV.Nodes(i).Expanded = True
    Next
End Sub

Sub CollapseAll()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView

    Dim i           As Long
    For i = 1 To TV.Nodes.count
        TV.Nodes(i).Expanded = False
    Next
End Sub

Sub clear()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Dim i           As Long
    For i = TV.Nodes.count To 1 Step -1
        TV.Nodes.Remove i
    Next
End Sub
Public Sub ChildrenCheck(Parent As MSComctlLib.Node, check As Boolean)
    ' In userform:
    ' Sub treeview1_NodeCheck(ByVal node As MSComctlLib.Node)
    '     TreeviewNodeCheck node, node.Checked
    ' End Sub
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView

    Dim child       As MSComctlLib.Node
    Parent.Checked = check

    ' Recursively check child nodes
    Set child = Parent.child
    While Not child Is Nothing
        ChildrenCheck child, check
        Set child = child.Next
    Wend
End Sub

Public Function GetLevel(ByVal theNode As Node) As Integer
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    GetLevel = 1
    Do Until theNode.Root = theNode.FirstSibling
        GetLevel = GetLevel + 1
        Set theNode = theNode.Parent
    Loop
End Function

Sub LoadRange( _
        FirstCell As Range, _
        ClearPreviousNodes As Boolean, _
        Expanded As Boolean)
    Rem example use
    Rem    LoadRange me.Treeview1,thisworkbook.sheets("TreeviewHierarchy").range("A1"),true,false
    Rem example of sheet structure
    Rem  |1|2|3
    Rem 1|A| |
    Rem 2| |1|
    Rem 3| | |1.1
    Rem 4| | |1.2
    Rem 5|B| |
    Rem 6| |2|
    Rem 7| | |2.1
    Rem 8| | |2.2
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Dim nP          As Node
    Dim c           As Excel.Range
    On Error Resume Next
    Dim childRange  As Range
    Set childRange = FirstCell.CurrentRegion.offset(0, 1)
    Set childRange = childRange.Resize(, childRange.Columns.count - 1)
    Set childRange = childRange.SpecialCells(xlCellTypeConstants)
    With TV
        If ClearPreviousNodes = True Then .Nodes.clear
        For Each c In FirstCell.Parent.Columns(FirstCell.Column).SpecialCells(xlCellTypeConstants)
            Set nP = .Nodes.Add(, , c.Address, c.value)
        Next
        For Each c In childRange
            If c.offset(, -1) = "" Then  'c.Value <> vbNullString And c.Address <> FirstCell.Address And
                Set nP = .Nodes(c.offset(, -1).End(xlUp).Address)
                If nP Is Nothing Then
                    MsgBox "ERROR: Parent node " & c.offset(, -1).End(xlUp).value & " not found...", vbExclamation, "Error"
                    Exit Sub
                End If
                .Nodes.Add nP, tvwChild, c.Address, c.value
                If Err.Number <> 0 Then
                    MsgBox "ERROR: The node " & c.value & " is a duplicate. All node descrptions must be unique", vbExclamation, "Error"
                    Exit Sub
                End If
                nP.Expanded = Expanded
            End If
        Next
        With .Nodes(FirstCell.Address)
            .Selected = True
            .EnsureVisible
        End With
    End With
End Sub

Public Sub LoadTreeArray( _
        arr, _
        ClearPreviousNodes As Boolean, _
        Expanded As Boolean)
    '@AssignedModule aTreeView
    '@INCLUDE PROCEDURE ArrayToRange2D
    '@INCLUDE PROCEDURE CreateOrSetSheet
    '@INCLUDE CLASS aTreeView

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim tmpSheet    As Worksheet
    Set tmpSheet = CreateOrSetSheet("TempOps", ThisWorkbook)
    tmpSheet.Cells.clear
    ArrayToRange2D arr, tmpSheet.Cells(1, 1)
    LoadRange tmpSheet.Cells(1, 1), ClearPreviousNodes, Expanded
    tmpSheet.Delete

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub LoadVBProjects()
    '@AssignedModule aTreeView
    '@INCLUDE PROCEDURE ProceduresOfModule
    '@INCLUDE PROCEDURE WorkbookProjectProtected
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aTreeView
    Me.clear
    Dim nP          As Node
    Dim nM          As Node
    Dim nS          As Node
    Dim x, y, s, p  As Variant
    Dim Module      As VBComponent
    On Error Resume Next
    Dim element
    For Each x In Array(Workbooks, AddIns)
        For Each y In x
            If Not WorkbookProjectProtected(Workbooks(y.Name)) Then
                If Err.Number = 0 Then
                    Set nP = TV.Nodes.Add(, , , y.Name)
                    For Each element In Array(vbext_ct_Document, vbext_ct_MSForm, vbext_ct_StdModule, vbext_ct_ClassModule)
                        For Each Module In Workbooks(y.Name).VBProject.VBComponents
                            If Module.Type = element Then
                                Set nM = TV.Nodes.Add(nP, tvwChild, , aModule.Initialize(Module).Name)
                                For Each p In ProceduresOfModule(Module)
                                    Set nS = TV.Nodes.Add(nM, tvwChild, , CStr(p))
                                Next p
                            End If
                        Next Module
                    Next
                End If
            End If
            Err.clear
        Next y
    Next x
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 10:07    Alex                select search item (aTreeView.cls > ActivateProjectElement)

Sub ActivateProjectElement()
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aTreeView
    '@INCLUDE USERFORM uCodeFinder
    Dim TargetWorkbook As Workbook

    Dim Module      As VBComponent
    Dim i           As Long
    Select Case GetLevel(TV.SelectedItem)
        Case Is = 1
        Case Is = 2
            With TV.SelectedItem
                Set TargetWorkbook = Workbooks(.Parent.text)
                On Error Resume Next
                Set Module = Workbooks(.Parent.text).VBProject.VBComponents(.text)
                On Error GoTo 0
                If Module Is Nothing Then Set Module = TargetWorkbook.VBProject.VBComponents(Workbooks(.Parent.text).Sheets(.text).codeName)
                aModule.Initialize(Module).Activate
            End With
        Case Is = 3
            With TV.SelectedItem
                Set TargetWorkbook = Workbooks(.Parent.Parent.text)
                On Error Resume Next
                Set Module = TargetWorkbook.VBProject.VBComponents(.Parent.text)
                On Error GoTo 0
                If Module Is Nothing Then
                    'get worksheet module instead
                    Set Module = TargetWorkbook.VBProject.VBComponents(TargetWorkbook.Sheets(.Parent.text).codeName)
                End If
                aModule.Initialize(Module).Activate
                For i = 1 To Module.CodeModule.countOfLines
                    If InStr(1, Module.CodeModule.lines(i, 1), "Sub " & .text) > 0 Or _
                            InStr(1, Module.CodeModule.lines(i, 1), "Function " & .text) > 0 Then
                        Module.CodeModule.CodePane.SetSelection i, 1, i, 1
                        Exit Sub
                    End If
                Next
            End With
        Case Is = 4
            With TV.SelectedItem
                Set TargetWorkbook = Workbooks(.Parent.Parent.Parent.text)
                Set Module = TargetWorkbook.VBProject.VBComponents(.Parent.Parent.text)
                Dim Procedure As String
                Procedure = .Parent.text
                Dim ap As New aProcedure
                ap.Initialize TargetWorkbook, Module, Procedure
                ap.Activate

                For i = ap.lines.FirstOfDeclaration To ap.lines.last
                    Dim StartPos As Long
                    StartPos = InStr(1, Module.CodeModule.lines(i, 1), uCodeFinder.TextBox1.value, vbTextCompare)
                    If StartPos > 0 Then
                        Dim endPos As Long: endPos = StartPos + Len(uCodeFinder.TextBox1.value)
                        Module.CodeModule.CodePane.SetSelection i, StartPos, i, endPos
                        Exit Sub
                    End If
                Next
            End With
    End Select
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 12:48    Alex                changed unnecessary condition check to removing childless nodes (aTreeView.cls > FindCodeEverywhere)

Public Sub FindCodeEverywhere(F As String)
    '@LastModified 2308181248
    '@AssignedModule aTreeView
    '@INCLUDE PROCEDURE WorkbookProjectProtected
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    '@INCLUDE CLASS aTreeView
    Dim nP          As Node
    Dim nM          As Node
    Dim nS          As Node
    Dim nF          As Node
    Dim x, y, s, p  As Variant
    Dim Module      As VBComponent
    Dim am          As aModule
    Dim ap          As aProcedure
    Dim aw          As aProject
    Dim i           As Long
    On Error Resume Next
    For Each x In Array(Workbooks, AddIns)
        For Each y In x
            If Not WorkbookProjectProtected(Workbooks(y.Name)) Then
                If Err.Number = 0 Then
                    Set aw = aProject.Initialize(Workbooks(y.Name))
                    'Dim ub As Long
                    'ub = UBound(Filter(Split(aw.Code, vbNewLine), F, True, vbTextCompare))
                    'Dim proceed As Boolean: proceed = InStr(1, aw.Code, F, vbTextCompare) > 0 'ub > -1 '
                    'If proceed Then
                    Set nP = TV.Nodes.Add(, , , y.Name)
                    For Each am In aModules.Initialize(aw.Project).Items
                        If am.Contains(F, False, False, True) Then
                            Set nM = TV.Nodes.Add(nP, tvwChild, , am.Name)
                            For Each ap In am.Procedures
                                If UBound(Filter(Split(ap.code.All, vbNewLine), F, True, vbTextCompare)) > -1 Then
                                    Set nS = TV.Nodes.Add(nM, tvwChild, , CStr(ap.Name))
                                    s = Filter(Split(ap.code.All, vbNewLine), F, True, vbTextCompare)
                                    For i = 0 To UBound(s)
                                        Set nF = TV.Nodes.Add(nS, tvwChild, , Trim(s(i)))
                                    Next i
                                End If
                            Next ap
                        End If
                    Next am
                    '                    End If
                    If nP.Children = 0 Then TV.Nodes.Remove nP.index
                End If
            End If
            Err.clear
        Next y
    Next x
End Sub

Sub ImageListLoadProjectIcons(imgList As ImageList)
    '@AssignedModule aTreeView
    '@INCLUDE CLASS aTreeView
    Dim strPath     As String
    strPath = ThisWorkbook.path & "\Lib\img\treeviewicons\"    ' "C:\Users\acer\Dropbox\SOFTWARE\EXCEL\0 Alex\treeviewicons\"
    With imgList.ListImages
        .Add , "Project", LoadPicture(strPath & "Project.jpg")
        .Add , "Sheet", LoadPicture(strPath & "Sheet.jpg")
        .Add , "Form", LoadPicture(strPath & "Form.jpg")
        .Add , "Module", LoadPicture(strPath & "Module.jpg")
        .Add , "Class", LoadPicture(strPath & "Class.jpg")
        .Add , "Macro", LoadPicture(strPath & "Macro.jpg")
        .Add , "Text", LoadPicture(strPath & "Text.jpg")
    End With
    TV.ImageList = imgList
End Sub

Sub TreeviewAssignProjectImages()
    '@AssignedModule aTreeView
    '@INCLUDE PROCEDURE ModuleTypeToString
    '@INCLUDE PROCEDURE ModuleOfWorksheet
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aTreeView
    Dim sModuleName As String
    Dim i           As Long
    Dim Module      As VBComponent
    Dim TargetWorkbook As Workbook
    For i = 1 To TV.Nodes.count
        Select Case GetLevel(TV.Nodes.Item(i))
            Case 1
                If InStr(1, TV.Nodes.Item(i).text, ".") = 0 Then GoTo SKIP
                TV.Nodes.Item(i).Image = 1    'tvImages.tvProject
            Case 2
                Set TargetWorkbook = Workbooks(TV.Nodes.Item(i).Parent.text)
                If InStr(1, TargetWorkbook.Name, ".") = 0 Then GoTo SKIP
                sModuleName = TV.Nodes.Item(i).text
                Set Module = Nothing
                On Error Resume Next
                Set Module = TargetWorkbook.VBProject.VBComponents(sModuleName)
                On Error GoTo 0
                If Module Is Nothing Then
                    Set Module = ModuleOfWorksheet(TargetWorkbook.Worksheets(TV.Nodes.Item(i).text))
                End If
                Select Case aModule.Initialize(Module).TypeToString
                    Case "Document"
                        TV.Nodes.Item(i).Image = 2    'tvImages.tvSheet
                    Case "UserForm"
                        TV.Nodes.Item(i).Image = 3    'tvImages.tvForm
                    Case "Module"
                        TV.Nodes.Item(i).Image = 4    'tvImages.tvModule
                    Case "Class"
                        TV.Nodes.Item(i).Image = 5    'tvImages.tvClass
                End Select
            Case 3
                TV.Nodes.Item(i).Image = 6    'tvImages.tvMacro
            Case 4
                TV.Nodes.Item(i).Image = 7    'tvImages.tvText
        End Select
SKIP:
    Next i
End Sub


