Attribute VB_Name = "M_BLOG"
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun FormatFolderToString  As String   :
' Fun myClassNames          As Variant  :
' Sub myClassesTOC                      :
' Sub myClassesToObsidian               :
'--------------------------------------------------
'@EndFolder ListOfProcedures
Option Explicit

Sub myClassesToObsidian()
    Dim regexPattern As String
    regexPattern = "\b(" & Join(myClassNames, "|") & ")\b"
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = regexPattern
    
    Dim targetFolder As String: targetFolder = "C:\Users\acer\Dropbox\Personal\Obsidian\Vault\Coding\VBA\Projects\VBIDE Extensibility\"
    Dim targetFile As String
    Dim m As aModule
    Dim i, index
    Dim out As String, TOC As String
    Dim it
    For Each it In myClassNames
        out = ""
        targetFile = targetFolder & it & ".md"
        Set m = aModule.Initialize(ThisWorkbook.VBProject.VBComponents(it))
        index = m.LineLike("'@EndFolder ListOfProcedures", False, True)
        If index > 0 Then
            Dim listOfProcs As String: listOfProcs = FormatFolderToString(m)
            out = "# ListOfProcedures" & vbNewLine & vbNewLine & listOfProcs
            TOC = TOC & vbNewLine & vbNewLine & "# [[" & it & "]]" & vbNewLine & vbNewLine & listOfProcs
        Else
            out = m.code
        End If
        
        out = regex.Replace(out, "[[$1]]")

        out = out & vbNewLine & vbNewLine & "# Code" & vbNewLine & vbNewLine & "```vb"
        out = out & vbNewLine & IIf(index > 0, Split(m.code, "'@EndFolder ListOfProcedures")(1), m.code)
        out = out & vbNewLine & "```"
        
        TxtOverwrite targetFile, out
    Next
    TxtOverwrite targetFolder & "0 - TOC.md", TOC
End Sub

Function myClassNames()
    myClassNames = Split("aCodeModule,aColorScheme,aComboBox,aDesigner,aFrame,aListBox,aListView,aModule,aModuleEnumItem,aModuleEnums,aModuleFolders,aModuleProcedures,aModules,aModuleTypeItem,aModuleTypes,aMultiPage,aProcedure,aProcedureArguments,aProcedureArgumentsItem,aProcedureCode,aProcedureCustomProperties,aProcedureDependencies,aProcedureFolder,aProcedureFormat,aProcedureInject,aProcedureLines,aProcedureMove,aProcedureScope,aProcedureVariables,aProcedureVariablesItem,aProject,aProjectDeclarations,aProjectReferences,aTreeView,aUserform", ",")
End Function

Function FormatFolderToString(m As aModule) As String
    Dim content As String
    content = m.Folders.ToString("ListOfProcedures")
    content = Replace(content, "'@FOLDER ListOfProcedures", "")
    content = Replace(content, "'@EndFolder ListOfProcedures", "")
    content = Replace(content, "' ", "")
    content = Replace(content, "'", "")
    content = Replace(content, ":", "")
    content = Replace(content, String(50, "-"), "")
    
    Dim arr As Variant
    arr = Split(content, vbNewLine)
    arr = cleanArray(arr)
    Dim Output
    ReDim Output(0 To UBound(arr) + 2, 0 To 2)
    Output(0, 0) = "| Type": Output(0, 1) = "| Procedure": Output(0, 2) = "| Returns"

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Output(i + 2, 0) = "| " & Split(arr(i), " ")(0)
        Output(i + 2, 1) = "| " & Split(arr(i), " ")(1)
        Output(i + 2, 2) = "|"
        If InStr(1, arr(i), " As ") > 0 Then
            Output(i + 2, 2) = Output(i + 2, 2) & Mid(arr(i), InStr(1, arr(i), " As "))
        Else
            Output(i + 2, 2) = Output(i + 2, 2)
        End If
    Next

    Dim ll As Long, x As Long, y As Long
    
    For y = LBound(Output, 2) To UBound(Output, 2)
        ' Calculate ll for the current column
        arr = ArraySubset2d(Output, , y, , 1)
        ll = LargestLength(arr) + 1
        Output(1, y) = "| " & String(ll - 3, "-")
        For x = LBound(Output, 1) To UBound(Output, 1)
            Output(x, y) = PadRight(Output(x, y), ll, Space(1))
            If y = UBound(Output, 2) Then Output(x, y) = Output(x, y) & " |"
        Next
    Next
    
    content = ""
    For i = LBound(Output) To UBound(Output)
        content = content & IIf(content <> "", vbNewLine, "")
        Dim element
        For Each element In ArraySubset2d(Output, i, , 1)
            content = content & element
        Next
    Next

    FormatFolderToString = content
End Function


Sub myClassesTOC()
    Dim m As aModule
    Dim it
    Dim content As String
    For Each it In myClassNames
      Set m = aModule.Initialize(ThisWorkbook.VBProject.VBComponents(it))
      content = vbNewLine & content & vbNewLine & "# " & it & vbNewLine
      content = content & vbNewLine & FormatFolderToString(m)
    Next
    
    CLIP content
End Sub
