Attribute VB_Name = "Dependencies"
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun ActiveCodepaneWorkbook               As Workbook     :
' Fun ActiveModule                         As VBComponent  :
' Fun ActiveProcedure                      As String       :
' Fun ActiveProject                        As VBProject    :
' Fun ArrayAllocated                       As Boolean      :
' Fun ArrayAppend                          As Variant      :
' Fun ArrayDimensionLength                 As Integer      :
' Fun ArrayDuplicatesRemove                As Variant      :
' Fun ArrayToCollection                    As Collection   :
' Fun AssignCPSvariables                   As Boolean      :
' Fun AssignModuleVariable                 As Boolean      :
' Fun AssignProcedureVariable              As Boolean      :
' Fun AssignWorkbookVariable               As Boolean      :
' Fun CLIP                                 As String       :
' Fun CheckPath                            As String       :
' Fun ClassNames                           As Variant      :
' Fun CleanTrim                            As String       :
' Fun CodepaneSelection                    As String       :
' Fun CollectionContains                   As Boolean      :
' Fun CollectionSort                       As Collection   :
' Fun CollectionsToArray2D                 As Variant      :
' Fun ComponentNames                       As Variant      :
' Fun DeclarationsKeywordSubstring         As String       :
' Fun DeclarationsTableKeywords            As Collection   :
' Fun DeclarationsWorksheetCreate          As Boolean      :
' Fun DevInfo                              As String       :
' Fun FileExists                           As Boolean      :
' Fun FolderExists                         As Boolean      :
' Fun FormatVBA7                           As String       :
' Fun GetMotherBoardProp                   As String       :
' Fun GetSheetByCodeName                   As Worksheet    :
' Fun IncludeNotImported                   As Variant      :
' Fun Len2                                 As Integer      :
' Fun LinkedClasses                        As Collection   :
' Fun LinkedDeclarations                   As Collection   :
' Fun LinkedProcedures                     As Collection   :
' Fun LinkedProceduresDeep                 As Collection   :
' Fun LinkedUserforms                      As Collection   :
' Fun ListOfIncludes                       As Variant      :
' Fun ModuleAddOrSet                       As VBComponent  :
' Fun ModuleCode                           As String       :
' Fun ModuleExists                         As Boolean      :
' Fun ModuleOfProcedure                    As VBComponent  :
' Fun ModuleOrSheetName                    As String       :
' Fun ModuleTypeToString                   As String       :
' Fun ProcedureAssignedModule              As VBComponent  :
' Fun ProcedureBodyLineFirst               As Long         :
' Fun ProcedureBodyLineFirstAfterComments  As Long         :
' Fun ProcedureCode                        As String       :
' Fun ProcedureExists                      As Boolean      :
' Fun ProcedureLastModAdd                  As Variant      :
' Fun ProcedureLastModified                As Variant      :
' Fun ProcedureLineContaining              As Long         :
' Fun ProcedureLinesCount                  As Long         :
' Fun ProcedureLinesFirst                  As Long         :
' Fun ProcedureLinesLast                   As Long         :
' Fun ProcedureTitle                       As String       :
' Fun ProcedureTitleLineCount              As Long         :
' Fun ProcedureTitleLineFirst              As Long         :
' Fun ProcedureTitleLineLast               As Long         :
' Fun ProceduresNotExported                As Variant      :
' Fun ProceduresOfModule                   As Collection   :
' Fun ProceduresOfTXT                      As Collection   :
' Fun ProceduresOfWorkbook                 As Collection   :
' Fun RegexTest                            As Boolean      :
' Fun StringLastModified                   As Variant      :
' Fun StringProcedureAssignedModule        As String       :
' Fun TXTReadFromUrl                       As String       :
' Fun ThisProject                          As Variant      :
' Fun TxtRead                              As String       :
' Fun URLExists                            As Boolean      :
' Fun UserformNames                        As Variant      :
' Fun WorkbookCode                         As String       :
' Fun WorkbookOfModule                     As Workbook     :
' Fun WorkbookOfProject                    As Workbook     :
' Fun WorksheetExists                      As Boolean      :
' Fun classCallsOfModule                   As Variant      :
' Fun cleanArray                           As Variant()    :
' Fun collectionToString                   As String       :
' Fun getDeclarations                      As Collection   :
' Fun getLastCell                          As Variant      :
' Fun getLastRow                           As Variant      :
' Fun vbarcFolders                         As Collection   :
' Sub AddLinkedLists                                       :
' Sub AddLinkedListsToActiveProcedure                      :
' Sub AddLinkedListsToActiveWorkbook                       :
' Sub AddLinkedListsToAllProcedures                        :
' Sub AddLinkedListsToProceduresOfModule                   :
' Sub AddListOfLinkedClassesToProcedure                    :
' Sub AddListOfLinkedDeclarationsToProcedure               :
' Sub AddListOfLinkedProceduresToProcedure                 :
' Sub AddListOfLinkedUserformsToProcedure                  :
' Sub ArrayQuickSort                                       :
' Sub ArrayToRange2D                                       :
' Sub CallSeparateProcedures                               :
' Sub CallTxtPrependContainedProcedures                    :
' Sub ClearClipboard                                       :
' Sub DeclarationsTableCreate                              :
' Sub DeclarationsTableSort                                :
' Sub ExportActiveProcedure                                :
' Sub ExportAllProcedures                                  :
' Sub ExportAllProceduresOfThisWorkbook                    :
' Sub ExportLinkedDeclaration                              :
' Sub ExportProcedure                                      :
' Sub ExportTargetProcedure                                :
' Sub FoldersCreate                                        :
' Sub FollowLink                                           :
' Sub ImportActiveProcedureDependencies                    :
' Sub ImportClass                                          :
' Sub ImportDeclaration                                    :
' Sub ImportProcedure                                      :
' Sub ImportProcedureDependencies                          :
' Sub ImportUserform                                       :
' Sub LinkedProceduresMoveHere                             :
' Sub ProcedureAssignedModuleAdd                           :
' Sub ProcedureLinesRemove                                 :
' Sub ProcedureLinesRemoveInclude                          :
' Sub ProcedureMoveHere                                    :
' Sub ProcedureMoveToAssignedModule                        :
' Sub ProcedureMoveToModule                                :
' Sub ProcedureReplace                                     :
' Sub ProjetFoldersCreate                                  :
' Sub RemoveComments                                       :
' Sub TxtOverwrite                                         :
' Sub TxtPrepend                                           :
' Sub TxtPrependContainedProcedures                        :
' Sub TxtSeparateProcedures                                :
'--------------------------------------------------
'@EndFolder ListOfProcedures


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Dependencies
'* Purpose    : list dependencies, export/import procedures/components
'* Copyright  :
'*
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : some time ago       Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


#If VBA7 Then
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#Else
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#End If

Rem ___CHANGE THESE TO MATCH YOUR FOLDER AND REPO____
'------------------------------------------------------------------------------
Public Const GITHUB_LIBRARY = "https://raw.githubusercontent.com/alexofrhodes/VBA-Library/"
'------------------------------------------------------------------------------
Public Const GITHUB_LIBRARY_DECLARATIONS = GITHUB_LIBRARY & "Declarations/"
Public Const GITHUB_LIBRARY_PROCEDURES = GITHUB_LIBRARY & "Procedures/"
Public Const GITHUB_LIBRARY_USERFORMS = GITHUB_LIBRARY & "Userforms/"
Public Const GITHUB_LIBRARY_CLASSES = GITHUB_LIBRARY & "Classes/"
'------------------------------------------------------------------------------
Public Const GITHUB_BLOG = "https://alexofrhodes.github.io/"
Public Const GITHUB_URL = "https://github.com/alexofrhodes/"
'------------------------------------------------------------------------------
Public Const AUTHOR_YOUTUBE = "https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg"
Public Const AUTHOR_VK = "https://vk.com/video/playlist/735281600_1"
Public Const AUTHOR_NAME = "Anastasiou Alex"
Public Const AUTHOR_EMAIL = "AnastasiouAlex@gmail.com"
Public Const AUTHOR_COPYRIGHT = ""
Public Const AUTHOR_OTHERTEXT = ""

Public Const GUID = "NBGD41100771701DDE7600"

Public ShowInVBE    As Boolean

'------------------------------------------------------------------------------
Public Function LOCAL_LIBRARY(): LOCAL_LIBRARY = LIBRARY_FOLDER: End Function
'------------------------------------------------------------------------------
Public Function LOCAL_LIBRARY_DECLARATIONS(): LOCAL_LIBRARY_DECLARATIONS = LOCAL_LIBRARY & "Declarations\": End Function
Public Function LOCAL_LIBRARY_PROCEDURES(): LOCAL_LIBRARY_PROCEDURES = LOCAL_LIBRARY & "Procedures\": End Function
Public Function LOCAL_LIBRARY_USERFORMS(): LOCAL_LIBRARY_USERFORMS = LOCAL_LIBRARY & "Userforms\": End Function
Public Function LOCAL_LIBRARY_CLASSES(): LOCAL_LIBRARY_CLASSES = LOCAL_LIBRARY & "Classes\": End Function

Function LIBRARY_FOLDER() As String
    If GetMotherBoardProp = GUID Then
        LIBRARY_FOLDER = "C:\Users\acer\Documents\GitHub\VBA-Library\"
    Else
        LIBRARY_FOLDER = Environ$("USERPROFILE") & "\Documents\vbArc\Library\"
    End If
End Function

Public Function AUTHOR_MEDIA() As String
    AUTHOR_MEDIA = "'* BLOG       : " & GITHUB_BLOG & vbNewLine & _
                   "'* GITHUB     : " & GITHUB_URL & vbNewLine & _
                   "'* YOUTUBE    : " & AUTHOR_YOUTUBE & vbNewLine & _
                   "'* VK         : " & AUTHOR_VK & vbNewLine & "'*" & vbNewLine
End Function

Function DevInfo() As String
    DevInfo = Join( _
                Array( _
                    "AUTHOR     " & AUTHOR_NAME, _
                    "EMAIL      " & AUTHOR_EMAIL, _
                    "BLOG       " & GITHUB_BLOG, _
                    "GITHUB     " & GITHUB_URL, _
                    "YOUTUBE    " & AUTHOR_YOUTUBE, _
                    "VK         " & AUTHOR_VK), _
                vbNewLine)
End Function


'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 08:29    Alex                (z_zTest.bas > ListOfIncludes)

Function ListOfIncludes(TargetWorkbook As Workbook)
'@LastModified 2308220829
'@INCLUDE CLASS aWorkbook
'@INCLUDE PROCEDURE ArrayQuickSort
'@INCLUDE PROCEDURE ArrayDuplicatesRemove
'@INCLUDE PROCEDURE ArrayReplace
'@INCLUDE PROCEDURE ArrayTrim
    Dim arr
    arr = ArrayTrim( _
            Split( _
                aProject.Initialize(ThisWorkbook).code, _
                vbNewLine))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = WorksheetFunction.Proper(arr(i))
    Next
    arr = ArrayDuplicatesRemove( _
            Filter( _
                Filter( _
                    arr, _
                    "'@INCLUDE PROCEDURE ", _
                    True, _
                    vbTextCompare), _
                """", _
                False))
    ArrayReplace arr, "'@INCLUDE PROCEDURE ", ""
    ArrayQuickSort arr
    
    ListOfIncludes = arr
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 08:28    Alex                (Dependencies.bas > ProceduresNotExported)

Function ProceduresNotExported(TargetWorkbook As Workbook)
'@LastModified 2308220828
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE ListOfIncludes
    Dim out
    Dim targetFile
    Dim Procedure
    For Each Procedure In ListOfIncludes(TargetWorkbook)
        targetFile = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"
        If Not FileExists(targetFile) Then
            out = out & IIf(out <> "", vbNewLine, "") & Procedure
        End If
    Next
    out = Split(out, vbNewLine)
    ProceduresNotExported = out
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 08:28    Alex                (Dependencies.bas > IncludeNotImported)

Function IncludeNotImported(TargetWorkbook As Workbook)
'@LastModified 2308220828
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE ListOfIncludes
    Dim out
    Dim targetFile
    Dim Procedure
    For Each Procedure In ListOfIncludes(TargetWorkbook)
        If Not ProcedureExists(TargetWorkbook, Procedure) Then
            out = out & IIf(out <> "", vbNewLine, "") & Procedure
        End If
    Next
    out = Split(out, vbNewLine)
    IncludeNotImported = out
End Function

'--------------------------------------------
Sub AddLinkedListsToActiveProcedure()
    AddLinkedLists ThisWorkbook, ActiveModule, ActiveProcedure
End Sub

Sub ExportActiveProcedure()
    ExportProcedure ThisWorkbook, ActiveModule, ActiveProcedure, ExportMergedTxt:=True
End Sub

Sub ExportAllProceduresOfThisWorkbook()
    ExportAllProcedures ThisWorkbook
End Sub

Sub ImportActiveProcedureDependencies()
    ImportProcedureDependencies ActiveProcedure, ThisWorkbook, ActiveModule, Overwrite:=True
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 30-10-2023 12:23    Alex                (Dependencies.bas > AddLinkedListsToAllProcedures : )

Sub AddLinkedListsToAllProcedures(TargetWorkbook As Workbook)
'@LastModified 2310301223
    Dim Procedure
    Dim Module      As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule And Module.Name <> "Dependencies" Then
            For Each Procedure In ProceduresOfModule(Module)
                AddLinkedLists TargetWorkbook, Module, CStr(Procedure)
            Next Procedure
        End If
    Next Module
    Toast "Done"
End Sub

Sub AddLinkedListsToActiveWorkbook()
    AddLinkedListsToAllProcedures ActiveCodepaneWorkbook
End Sub

Sub AddLinkedListsToProceduresOfModule(Module As VBComponent)
    Dim Procedure
    On Error GoTo EH
    For Each Procedure In ProceduresOfModule(Module)
        Debug.Print Procedure
        AddLinkedLists , Module, CStr(Procedure)
    Next Procedure
    Debug.Print vbNewLine & "---" & "Done"
    Exit Sub
EH:
    Debug.Print "Error at: " & Module.Name & vbTab & Procedure
    Resume Next
End Sub

Sub ExportAllProcedures(TargetWorkbook As Workbook)
    Dim Procedure
    Dim Module      As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then
            For Each Procedure In ProceduresOfModule(Module)
                ExportProcedure TargetWorkbook, Module, CStr(Procedure), False
            Next Procedure
        End If
    Next Module
End Sub

Sub RemoveComments(TargetWorkbook As Workbook)
    Dim Module      As VBComponent
    Dim s           As String
    Dim i           As Long
    For Each Module In TargetWorkbook.VBProject.VBComponents
        For i = Module.CodeModule.countOfLines To 1 Step -1
            s = Trim(Module.CodeModule.lines(i, 1))
            If s Like "'*" Then Module.CodeModule.DeleteLines i, 1
        Next i
    Next
End Sub

Function ArrayAppend(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    Dim holdarr     As Variant
    Dim ub1         As Long
    Dim ub2         As Long
    Dim i           As Long
    Dim newind      As Long
    If IsEmpty(arr1) Or Not IsArray(arr1) Then
        arr1 = Array()
    End If
    If IsEmpty(arr2) Or Not IsArray(arr2) Then
        arr2 = Array()
    End If
    ub1 = UBound(arr1)
    ub2 = UBound(arr2)
    If ub1 = -1 Then
        ArrayAppend = arr2
        Exit Function
    End If
    If ub2 = -1 Then
        ArrayAppend = arr1
        Exit Function
    End If
    holdarr = arr1
    ReDim Preserve holdarr(ub1 + ub2 + 1)
    newind = UBound(arr1) + 1
    For i = 0 To ub2
        If VarType(arr2(i)) = vbObject Then
            Set holdarr(newind) = arr2(i)
        Else
            holdarr(newind) = arr2(i)
        End If
        newind = newind + 1
    Next i
    ArrayAppend = holdarr
End Function

Public Sub ArrayQuickSort(ByRef SortableArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next
    Dim i           As Long
    Dim j           As Long
    Dim varMid      As Variant
    Dim varX        As Variant
    If IsEmpty(SortableArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortableArray), "()") < 1 Then
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortableArray)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortableArray)
    End If
    If lngMin >= lngMax Then
        Exit Sub
    End If
    i = lngMin
    j = lngMax
    varMid = Empty
    varMid = SortableArray((lngMin + lngMax) \ 2)
    If IsObject(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If
    While i <= j
        While SortableArray(i) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortableArray(j) And j > lngMin
            j = j - 1
        Wend
        If i <= j Then
            varX = SortableArray(i)
            SortableArray(i) = SortableArray(j)
            SortableArray(j) = varX
            i = i + 1
            j = j - 1
        End If
    Wend
    If (lngMin < j) Then Call ArrayQuickSort(SortableArray, lngMin, j)
    If (i < lngMax) Then Call ArrayQuickSort(SortableArray, i, lngMax)
End Sub

Public Function cleanArray(varArray As Variant) As Variant()
    Dim TempArray() As Variant
    Dim OldIndex    As Integer
    Dim NewIndex    As Integer
    Dim Output      As String
    If Not ArrayAllocated(varArray) Then Exit Function
    ReDim TempArray(LBound(varArray) To UBound(varArray))
    For OldIndex = LBound(varArray) To UBound(varArray)
        Output = CleanTrim(varArray(OldIndex))
        If Not Output = "" Then
            TempArray(NewIndex) = Output
            NewIndex = NewIndex + 1
        End If
    Next OldIndex
    ReDim Preserve TempArray(LBound(varArray) To NewIndex - 1)
    cleanArray = TempArray
End Function

Function ArrayDuplicatesRemove(myArray As Variant) As Variant
    Dim nFirst As Long, nLast As Long, i As Long
    Dim Item        As String

    Dim arrTemp()   As String
    Dim coll        As New Collection
    If Not ArrayAllocated(myArray) Then Exit Function
    nFirst = LBound(myArray)
    nLast = UBound(myArray)
    ReDim arrTemp(nFirst To nLast)

    For i = nFirst To nLast
        arrTemp(i) = CStr(myArray(i))
    Next i

    On Error Resume Next
    For i = nFirst To nLast
        coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.clear
    On Error GoTo 0

    nLast = coll.count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)

    For i = nFirst To nLast
        arrTemp(i) = coll(i - nFirst + 1)
    Next i

    ArrayDuplicatesRemove = arrTemp

End Function

Public Function ArrayToCollection(Items As Variant) As Collection
    If Not ArrayAllocated(Items) Then Exit Function
    Dim coll        As New Collection
    Dim i           As Integer
    For i = LBound(Items) To UBound(Items)
        coll.Add Items(i)
    Next
    Set ArrayToCollection = coll
End Function

Function CleanTrim(ByVal s As String, Optional ConvertNonBreakingSpace As Boolean = True) As String
    Dim x As Long, CodesToClean As Variant
    CodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
            21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
    If ConvertNonBreakingSpace Then s = Replace(s, Chr(160), " ")
    s = Replace(s, vbCr, "")
    For x = LBound(CodesToClean) To UBound(CodesToClean)
        If InStr(s, Chr(CodesToClean(x))) Then
            s = Replace(s, Chr(CodesToClean(x)), vbNullString)
        End If
    Next
    Do While InStr(1, s, "  ") > 0
        s = VBA.Replace(s, "  ", " ")
    Loop
    CleanTrim = s
    CleanTrim = Trim(s)
End Function

Sub AddLinkedLists(Optional TargetWorkbook As Workbook, _
                    Optional Module As VBComponent, _
                    Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Exit Sub
    ProcedureLinesRemoveInclude TargetWorkbook, Module, Procedure
    ProcedureAssignedModuleAdd TargetWorkbook, Module, Procedure
    AddListOfLinkedProceduresToProcedure TargetWorkbook, Module, Procedure
    AddListOfLinkedClassesToProcedure TargetWorkbook, Module, Procedure
    AddListOfLinkedUserformsToProcedure TargetWorkbook, Module, Procedure
    AddListOfLinkedDeclarationsToProcedure TargetWorkbook, Module, Procedure
End Sub


Sub AddListOfLinkedClassesToProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional ProcedureName As String)

    If Not AssignCPSvariables(TargetWorkbook, Module, ProcedureName) Then Stop
    Dim ListOfImports As String
    Dim code        As String
    code = ProcedureCode(TargetWorkbook, Module, ProcedureName)
    Dim myClasses   As Collection
    Set myClasses = LinkedClasses(TargetWorkbook, Module, ProcedureName)
    Dim element     As Variant
    For Each element In myClasses
        If InStr(1, code, "@INCLUDE CLASS " & element) = 0 _
                And InStr(1, ListOfImports, "@INCLUDE CLASS " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE CLASS " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE CLASS " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        Module.CodeModule.InsertLines _
                ProcedureBodyLineFirstAfterComments(Module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedDeclarationsToProcedure( _
                                            Optional TargetWorkbook As Workbook, _
                                            Optional Module As VBComponent, _
                                            Optional ProcedureName As String)

    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim ListOfImports As String
    If Module Is Nothing Then Set Module = ModuleOfProcedure(TargetWorkbook, ProcedureName)
    Dim ProcedureText As String
    ProcedureText = ProcedureCode(TargetWorkbook, Module, ProcedureName)
    Dim myDeclarations As Collection
    Set myDeclarations = LinkedDeclarations(TargetWorkbook, Module, ProcedureName)
    Dim coll        As New Collection
    Dim element     As Variant
    For Each element In myDeclarations
        If InStr(1, ProcedureText, "'@INCLUDE DECLARATION " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE DECLARATION " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE DECLARATION " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        Module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(Module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedProceduresToProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional ProcedureName As String)

    If Not AssignCPSvariables(TargetWorkbook, Module, ProcedureName) Then Stop
    Dim Procedures  As Collection
    Set Procedures = LinkedProcedures(TargetWorkbook, Module, ProcedureName)
    Dim ListOfImports As String
    Dim code        As String
    code = ProcedureCode(TargetWorkbook, Module, ProcedureName)
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If InStr(1, code, "@INCLUDE PROCEDURE " & Procedure) = 0 And InStr(1, ListOfImports, "@INCLUDE PROCEDURE " & Procedure) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE PROCEDURE " & Procedure
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE PROCEDURE " & Procedure
            End If
        End If
    Next
    If ListOfImports <> "" Then
        Module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(Module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedUserformsToProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional ProcedureName As String)

    If Not AssignCPSvariables(TargetWorkbook, Module, ProcedureName) Then Stop

    Dim ListOfImports As String
    Dim code        As String
    code = ProcedureCode(TargetWorkbook, Module, ProcedureName)
    Dim myClasses   As Collection
    Set myClasses = LinkedUserforms(TargetWorkbook, Module, ProcedureName)
    Dim element     As Variant
    For Each element In myClasses
        If InStr(1, code, "@INCLUDE USERFORM " & element) = 0 And InStr(1, ListOfImports, "@INCLUDE USERFORM " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE USERFORM " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE USERFORM " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        Module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(Module, ProcedureName), ListOfImports
    End If
End Sub

Public Function ActiveProcedure() As String
    Application.VBE.ActiveCodePane.GetSelection L1&, c1&, L2&, c2&
    ActiveProcedure = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(L1&, vbext_pk_Proc)
End Function

Public Function ActiveModule() As VBComponent
    '@LastModified 2308171258
    If Not Application.VBE.ActiveCodePane Is Nothing Then
        Set ActiveModule = Application.VBE.ActiveCodePane.CodeModule.Parent
    End If
    
'    Dim Module1 As VBComponent
'    'may erroneously return userform or worksheet
'
'    Set Module1 = Application.VBE.SelectedVBComponent
'    Dim Module2 As VBComponent
'    If Not Application.VBE.ActiveCodePane Is Nothing Then Set Module2 = Application.VBE.ActiveCodePane.CodeModule.Parent
'
'    If Module1.Name = Module2.Name Then
'        Set ActiveModule = Module1
'    Else
'        Dim ans As Long
'        ans = MsgBox("SelectedVBComponent <> ActiveCodePane.CodeModule.Parent" & vbLf & _
'                      "Choose " & Module1.Name & " ?    Click no to choose " & Module2.Name, _
'                      vbExclamation + vbYesNoCancel)
'        Select Case ans
'            Case vbYes: Set ActiveModule = Module1
'            Case vbNo: Set ActiveModule = Module2
'            Case vbCancel: Stop
'        End Select
'    End If
End Function

Public Function ThisProject()
'    Dim appName As String: appName = appName
'    #If appName = "Microsoft Word" Then
'        Set ThisProject = ThisDocument.VBProject
'    #ElseIf appName = "Microsoft Excel" Then
        Set ThisProject = ThisWorkbook.VBProject
'    #ElseIf appName = "Microsoft PowerPoint" Then
'
'    #ElseIf appName = "Microsoft Access" Then
'
'    #End If
End Function

Public Function ActiveProject() As VBProject
    Dim Module As VBComponent
    Set Module = ActiveModule
    If Not Module Is Nothing Then
        Set ActiveProject = Module.Collection.Parent
    End If
End Function

Public Function ActiveCodepaneWorkbook() As Workbook
    On Error GoTo ErrorHandler
    Dim WorkbookName As String
    WorkbookName = ActiveModule.Collection.Parent.fileName
    WorkbookName = Right(WorkbookName, Len(WorkbookName) - InStrRev(WorkbookName, "\"))
    Set ActiveCodepaneWorkbook = Workbooks(WorkbookName)
    Exit Function
ErrorHandler:
    MsgBox "doesn't work on new-unsaved workbooks"
End Function

Public Function ArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    ArrayAllocated = IsArray(arr) And (Not IsError(LBound(arr, 1))) And LBound(arr, 1) <= UBound(arr, 1)
End Function

Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    Dim i           As Integer
    Dim test        As Long
    On Error GoTo catch
    Do
        i = i + 1
        test = UBound(SourceArray, i)
    Loop
catch:
    ArrayDimensionLength = i - 1
End Function

Public Sub ArrayToRange2D(arr2D As Variant, cell As Range)

    If ArrayDimensionLength(arr2D) = 1 Then arr2D = WorksheetFunction.Transpose(arr2D)
    Dim dif         As Long
    dif = IIf(LBound(arr2D, 1) = 0, 1, 0)
    Dim rng         As Range
    Set rng = cell.Resize(UBound(arr2D, 1) + dif, UBound(arr2D, 2) + dif)

    If Application.WorksheetFunction.CountA(rng) > 0 Then
        Exit Sub
    End If

    rng.value = arr2D
End Sub

Function AssignCPSvariables( _
        ByRef TargetWorkbook As Workbook, _
        ByRef Module As VBComponent, _
        ByRef Procedure As String) As Boolean

    If Not AssignWorkbookVariable(TargetWorkbook) Then Exit Function
    If Not AssignModuleVariable(TargetWorkbook, Module, Procedure) Then Exit Function
    If Not AssignProcedureVariable(TargetWorkbook, Procedure) Then Exit Function
    AssignCPSvariables = True

End Function

Function AssignModuleVariable( _
        ByVal TargetWorkbook As Workbook, _
        ByRef Module As VBComponent, _
        Optional ByVal Procedure As String) As Boolean
    If Module Is Nothing Then
        If Procedure = "" Then
            Set Module = ActiveModule
        End If
        On Error Resume Next
        Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
        On Error GoTo 0
    End If
    AssignModuleVariable = Not Module Is Nothing
End Function

Function AssignProcedureVariable(TargetWorkbook As Workbook, ByRef Procedure As String) As Boolean
    If Procedure = "" Then
        Dim cps     As String
        cps = CodepaneSelection
        If Len(cps) > 0 Then
            Procedure = cps
        Else
            Procedure = ActiveProcedure
        End If
        If Not ProcedureExists(TargetWorkbook, Procedure) Then
            Debug.Print Procedure & " not found in Workbook " & TargetWorkbook.Name
        End If
    End If
    AssignProcedureVariable = Not Procedure = ""
End Function

Function AssignWorkbookVariable(ByRef TargetWorkbook As Workbook) As Boolean
    If TargetWorkbook Is Nothing Then
        On Error Resume Next
        Set TargetWorkbook = ActiveCodepaneWorkbook
        On Error GoTo 0
    End If
    AssignWorkbookVariable = Not TargetWorkbook Is Nothing
End Function

Function CheckPath(path) As String
    Dim RetVal
    RetVal = "I"
    If (RetVal = "I") And FileExists(path) Then RetVal = "F"
    If (RetVal = "I") And FolderExists(path) Then RetVal = "D"
    If (RetVal = "I") And URLExists(path) Then RetVal = "U"
    CheckPath = RetVal
End Function

Function ClassNames(Optional TargetWorkbook As Workbook)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Set ClassNames = ComponentNames(vbext_ct_ClassModule, TargetWorkbook)
End Function

Public Function CodepaneSelection() As String
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    If EndLine - StartLine = 0 Then
        CodepaneSelection = Mid(Application.VBE.ActiveCodePane.CodeModule.lines(StartLine, 1), StartColumn, EndColumn - StartColumn)
        Exit Function
    End If
    Dim str         As String
    Dim i           As Long
    For i = StartLine To EndLine
        If str = "" Then
            str = Mid(Application.VBE.ActiveCodePane.CodeModule.lines(i, 1), StartColumn)
        ElseIf i < EndLine Then
            str = str & vbNewLine & Application.VBE.ActiveCodePane.CodeModule.lines(i, 1)
        Else
            str = str & vbNewLine & Left(Application.VBE.ActiveCodePane.CodeModule.lines(i, 1), EndColumn - 1)
        End If
    Next
    CodepaneSelection = str
End Function

Public Function CollectionContains( _
        Kollection As Collection, _
        Optional Key As Variant, _
        Optional Item As Variant) As Boolean
    Dim strKey      As String
    Dim var         As Variant
    If Not IsMissing(Key) Then
        strKey = CStr(Key)
        On Error Resume Next
        CollectionContains = True
        var = Kollection(strKey)
        If Err.Number = 91 Then GoTo CheckForObject
        If Err.Number = 5 Then GoTo NotFound
        On Error GoTo 0
        Exit Function
CheckForObject:
        If IsObject(Kollection(strKey)) Then
            CollectionContains = True
            On Error GoTo 0
            Exit Function
        End If
NotFound:
        CollectionContains = False
        On Error GoTo 0
        Exit Function
    ElseIf Not IsMissing(Item) Then
        CollectionContains = False
        For Each var In Kollection
            If var = Item Then
                CollectionContains = True
                Exit Function
            End If
        Next var
    Else
        CollectionContains = False
    End If
End Function

Public Function CollectionSort(colInput As Collection) As Collection
    Dim iCounter    As Integer
    Dim iCounter2   As Integer
    Dim temp        As Variant
    Set CollectionSort = New Collection
    For iCounter = 1 To colInput.count - 1
        For iCounter2 = iCounter + 1 To colInput.count
            If colInput(iCounter) > colInput(iCounter2) Then
                temp = colInput(iCounter2)
                colInput.Remove iCounter2
                colInput.Add temp, , iCounter
            End If
        Next iCounter2
    Next iCounter
    Set CollectionSort = colInput
End Function

Function CollectionsToArray2D(collections As Collection) As Variant
    If collections.count = 0 Then Exit Function
    Dim columnCount As Long
    columnCount = collections.count
    Dim rowCount    As Long
    rowCount = collections.Item(1).count
    Dim var         As Variant
    ReDim var(1 To rowCount, 1 To columnCount)
    Dim cols        As Long
    Dim rows        As Long
    For rows = 1 To rowCount
        For cols = 1 To collections.count
            var(rows, cols) = collections(cols).Item(rows)
        Next cols
    Next rows
    CollectionsToArray2D = var
End Function

Function ComponentNames( _
        moduleType As vbext_ComponentType, _
        Optional TargetWorkbook As Workbook)
    Dim coll        As New Collection
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module      As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = moduleType Then
            coll.Add Module.Name
        End If
    Next
    Set ComponentNames = coll
End Function

Function DeclarationsKeywordSubstring(str As Variant, Optional delim As String _
        , Optional afterWord As String _
        , Optional beforeWord As String _
        , Optional counter As Integer _
        , Optional outer As Boolean _
        , Optional includeWords As Boolean) As String
    Dim i           As Long
    If afterWord = "" And beforeWord = "" And counter = 0 Then
        MsgBox ("Pass at least 1 parameter betweenn -AfterWord- , -BeforeWord- , -counter-")
        Exit Function
    End If
    If TypeName(str) = "String" Then
        If delim <> "" Then
            str = Split(str, delim)
            If UBound(str) <> 0 Then
                If afterWord = "" And beforeWord = "" And counter <> 0 Then
                    If counter - 1 <= UBound(str) Then
                        DeclarationsKeywordSubstring = str(counter - 1)
                        Exit Function
                    End If
                End If
                For i = LBound(str) To UBound(str)
                    If afterWord <> "" And beforeWord = "" Then
                        If i <> 0 Then
                            If str(i - 1) = afterWord Or str(i - 1) = "#" & afterWord Then
                                DeclarationsKeywordSubstring = str(i)
                                Exit Function
                            End If
                        End If
                    ElseIf afterWord = "" And beforeWord <> "" Then
                        If i <> UBound(str) Then
                            If str(i + 1) = beforeWord Or str(i + 1) = "#" & beforeWord Then
                                DeclarationsKeywordSubstring = str(i)
                                Exit Function
                            End If
                        End If
                    ElseIf afterWord <> "" And beforeWord <> "" Then
                        If i <> 0 And i <> UBound(str) Then
                            If (str(i - 1) = afterWord Or str(i - 1) = "#" & afterWord) And (str(i + 1) = beforeWord Or str(i + 1) = "#" & beforeWord) Then
                                DeclarationsKeywordSubstring = str(i)
                                Exit Function
                            End If
                        End If
                    End If
                Next i
            End If
        Else
            If InStr(1, str, afterWord) > 0 And InStr(1, str, beforeWord) > 0 Then
                If includeWords = False Then
                    DeclarationsKeywordSubstring = Mid(str, InStr(1, str, afterWord) + Len(afterWord))
                Else
                    DeclarationsKeywordSubstring = Mid(str, InStr(1, str, afterWord))
                End If
                If outer = True Then
                    If includeWords = False Then
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStrRev(DeclarationsKeywordSubstring, beforeWord) - 1)
                    Else
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStrRev(DeclarationsKeywordSubstring, beforeWord) + Len(beforeWord) - 1)
                    End If
                Else
                    If includeWords = False Then
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStr(1, DeclarationsKeywordSubstring, beforeWord) - 1)
                    Else
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStr(1, DeclarationsKeywordSubstring, beforeWord) + Len(beforeWord) - 1)
                    End If
                End If
                Exit Function
            End If
        End If
    Else
    End If
    DeclarationsKeywordSubstring = vbNullString
End Function

Sub DeclarationsTableCreate(TargetWorkbook As Workbook)

    DeclarationsWorksheetCreate

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    If Format(Now, "YYMMDDHHNN") - TargetWorksheet.Range("Z1").value < 60 Then Exit Sub

    TargetWorksheet.Range("A2").CurrentRegion.offset(1).clear
    ArrayToRange2D CollectionsToArray2D( _
            getDeclarations( _
            wb:=TargetWorkbook, _
            includeScope:=True, _
            includeType:=True, _
            includeKeywords:=True, _
            includeDeclarations:=True, _
            includeComponentName:=True, _
            includeComponentType:=True)), _
            TargetWorksheet.Range("A2")

    TargetWorksheet.Range("Z1").value = Format(Now, "YYMMDDHHNN")

    DeclarationsTableSort
End Sub


Function DeclarationsTableKeywords() As Collection
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    Dim Lr          As Long: Lr = getLastRow(TargetWorksheet)
    Dim coll        As New Collection
    Dim cell        As Range
    For Each cell In TargetWorksheet.Range("C2:C" & Lr)
        On Error Resume Next
        coll.Add cell.text, cell.text
        On Error GoTo 0
    Next
    Set DeclarationsTableKeywords = coll
End Function


Sub DeclarationsTableSort()

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Worksheets("Declarations_Table")

    Dim sort1       As String: sort1 = "B1"
    Dim sort2       As String: sort2 = "C1"
    Dim sort3       As String    ': sort3 = "D1"

    With TargetWorksheet.Sort
        .SortFields.clear
        .SortFields.Add Key:=TargetWorksheet.Range(sort1), Order:=xlAscending

        If Not sort2 = "" Then
            .SortFields.Add Key:=TargetWorksheet.Range(sort2), Order:=xlAscending
        End If
        If Not sort3 = "" Then
            .SortFields.Add Key:=TargetWorksheet.Range(sort3), Order:=xlAscending
        End If

        .SetRange TargetWorksheet.Range("A1").CurrentRegion
        .Header = xlYes
        .Apply
    End With

End Sub



Function DeclarationsWorksheetCreate() As Boolean
    If WorksheetExists("Declarations_Table", ThisWorkbook) Then Exit Function
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets.Add
    With TargetWorksheet
        .Name = "Declarations_Table"
        .Cells.VerticalAlignment = xlVAlignTop
        .Range("A1:F1").value = Split("SCOPE,TYPE,NAME,CODE,MODULE TYPE,MODULE NAME", ",")
        .rows(1).Cells.Font.Bold = True
        .rows(1).Cells.Font.Size = 14
    End With
End Function

Sub ExportLinkedDeclaration(TargetWorkbook As Workbook, DeclarationName As String)
    DeclarationsTableCreate TargetWorkbook
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")

    Dim codeName    As String
    Dim codeText    As String
    Dim cell        As Range
    On Error Resume Next
    Set cell = TargetWorksheet.Columns(3).Find(DeclarationName, LookAt:=xlWhole)
    On Error GoTo 0
    If cell Is Nothing Then Exit Sub

    codeName = DeclarationName
    codeText = cell.offset(0, 1).text
    TxtOverwrite LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt", codeText

End Sub



Sub ExportProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional ProcedureName As String, _
        Optional ExportMergedTxt As Boolean)

    If Not AssignCPSvariables(TargetWorkbook, Module, ProcedureName) Then Exit Sub

    ProjetFoldersCreate

    Dim ExportedProcedures As New Collection
    On Error GoTo ErrorHandler

    ExportedProcedures.Add CStr(ProcedureName), CStr(ProcedureName)

    Dim Procedure
    For Each Procedure In LinkedProceduresDeep(ProcedureName, TargetWorkbook)
        ExportedProcedures.Add CStr(Procedure), CStr(Procedure)
    Next

    If ExportedProcedures.count > 1 Then
        For Each Procedure In ExportedProcedures
            ExportTargetProcedure TargetWorkbook, , CStr(Procedure)
        Next
        If ExportMergedTxt Then
            Dim MergedName As String: MergedName = "Merged_" & ProcedureName
            Dim fileName As String: fileName = LOCAL_LIBRARY_PROCEDURES & MergedName & ".txt"
            Dim MergedString As String

            For Each Procedure In ExportedProcedures
                MergedString = MergedString & vbNewLine & ProcedureCode(TargetWorkbook, , Procedure)
            Next
            Debug.Print "OVERWROTE " & MergedName
            TxtOverwrite fileName, MergedString
            TxtPrependContainedProcedures fileName
        End If
    End If

    FollowLink LOCAL_LIBRARY_PROCEDURES

    Exit Sub
ErrorHandler:
    MsgBox "An error occured in Sub ExportProcedure"
End Sub

Sub ExportTargetProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As String)

    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Exit Sub

    Dim proclastmod
    proclastmod = ProcedureLastModified(TargetWorkbook, Module, Procedure)
    If proclastmod = 0 Then
        AddLinkedLists TargetWorkbook, Module, Procedure
        proclastmod = ProcedureLastModAdd(TargetWorkbook, Module, Procedure)
    End If

    Dim code        As String
    code = ProcedureCode(TargetWorkbook, Module, CStr(Procedure))
    Dim FileFullName As String
    FileFullName = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"
    If FileExists(FileFullName) Then
        Dim filelastmod
        filelastmod = StringLastModified(TxtRead(FileFullName))
        If proclastmod > filelastmod Then
            Debug.Print "OVERWROTE " & Procedure
            TxtOverwrite FileFullName, code
        End If
    Else
        Debug.Print "NEW " & Procedure
        TxtOverwrite FileFullName, code
    End If

    Dim element
    For Each element In LinkedUserforms(TargetWorkbook, Module, CStr(Procedure))
        TargetWorkbook.VBProject.VBComponents(element).Export LOCAL_LIBRARY_USERFORMS & element & ".frm"
    Next
    For Each element In LinkedClasses(TargetWorkbook, Module, CStr(Procedure))
        TargetWorkbook.VBProject.VBComponents(element).Export LOCAL_LIBRARY_CLASSES & element & ".cls"
    Next
    For Each element In LinkedDeclarations(TargetWorkbook, Module, CStr(Procedure))
        ExportLinkedDeclaration TargetWorkbook, CStr(element)
    Next
End Sub

Public Function FileExists(ByVal fileName As String) As Boolean
    If InStr(1, fileName, "\") = 0 Then Exit Function
    If Right(fileName, 1) = "\" Then fileName = Left(fileName, Len(fileName) - 1)
    FileExists = (Dir(fileName, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function

Function FolderExists(ByVal strPath As String) As Boolean
'@LastModified 2310251309
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Sub FoldersCreate(FolderPath As String)
    On Error Resume Next
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim ArrayElement As Variant
    individualFolders = Split(FolderPath, "\")
    For Each ArrayElement In individualFolders
        tempFolderPath = tempFolderPath & ArrayElement & "\"
        If FolderExists(tempFolderPath) = False Then
            MkDir tempFolderPath
        End If
    Next ArrayElement
End Sub

Sub FollowLink(FolderPath As String)
    If Right(FolderPath, 1) = "\" Then FolderPath = Left(FolderPath, Len(FolderPath) - 1)
    On Error Resume Next
    Dim oShell      As Object
    Dim Wnd         As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.document.Folder.Self.path = FolderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=FolderPath, NewWindow:=True
End Sub

Function FormatVBA7(str As String) As String
'FormatVBA7(join(filter(filter(split(aproject.initialize(thisworkbook).Code,vbnewline),"Declare ",True ,vbTextCompare),"""" & "Declare", False,vbTextCompare),vbnewline))
    Dim SelectedText
    SelectedText = str
    SelectedText = Replace(SelectedText, " _" & vbNewLine, "")
    SelectedText = Split(SelectedText, vbNewLine)
    Dim IsVba7      As String
    Dim NotVba7     As String
    Dim colIsVBA7   As New Collection
    Dim colNotVBA7  As New Collection
    Dim i           As Long
    For i = LBound(SelectedText) To UBound(SelectedText)
        If InStr(1, SelectedText(i), "PtrSafe", vbTextCompare) Then
            IsVba7 = SelectedText(i)
            NotVba7 = Replace(SelectedText(i), "Declare ptrsafe ", "Declare ", , , vbTextCompare)
        Else
            IsVba7 = Replace(SelectedText(i), "Declare ", "Declare PtrSafe ")
            NotVba7 = SelectedText(i)
        End If
        colIsVBA7.Add IsVba7
        colNotVBA7.Add NotVba7
    Next
    Set colIsVBA7 = CollectionSort(colIsVBA7)
    Set colNotVBA7 = CollectionSort(colNotVBA7)
    Dim out         As String
    out = "#If VBA7 then" & vbNewLine & _
            collectionToString(colIsVBA7, vbNewLine) & vbNewLine & _
            "#Else" & vbNewLine & _
            collectionToString(colNotVBA7, vbNewLine) & vbNewLine & _
            "#End If"
    FormatVBA7 = out

End Function

Function GetMotherBoardProp() As String

    Dim strComputer As String
    Dim objSvcs     As Object
    Dim objItms As Object, objItm As Object
    Dim vItem
    strComputer = "."
    Set objSvcs = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set objItms = objSvcs.execquery("Select * from Win32_BaseBoard")
    For Each objItm In objItms
        GetMotherBoardProp = objItm.SerialNumber
    Next

    Set objSvcs = Nothing
End Function

Public Function GetSheetByCodeName(wb As Workbook, codeName As String) As Worksheet
    Dim sh          As Worksheet
    For Each sh In wb.Worksheets
        If UCase(sh.codeName) = UCase(codeName) Then Set GetSheetByCodeName = sh: Exit For
    Next sh
End Function

Sub ImportClass( _
        Optional ClassName As String, _
        Optional TargetWorkbook As Workbook, _
        Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If ClassName = "" Then ClassName = CodepaneSelection
    If ClassName = "" Or InStr(1, ClassName, " ") > 0 Then Exit Sub
    Dim FilePath    As String
    FilePath = LOCAL_LIBRARY_CLASSES & ClassName & ".cls"
    If CheckPath(FilePath) = "I" Then
        On Error Resume Next
        Dim code    As String
        code = TXTReadFromUrl(GITHUB_LIBRARY_CLASSES & ClassName & ".cls")
        On Error GoTo 0
        If Len(code) > 0 And Not UCase(code) Like ("*NOT FOUND*") Then
            TxtOverwrite FilePath, code
        Else
            MsgBox "File " & ClassName & ".cls not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(ClassName, TargetWorkbook) Then
        If Overwrite = True Then
            TargetWorkbook.VBProject.VBComponents.Remove TargetWorkbook.VBProject.VBComponents(ClassName)
        Else
            Exit Sub
        End If
    End If
    TargetWorkbook.VBProject.VBComponents.Import FilePath
End Sub


Sub ImportDeclaration( _
        Optional DeclarationName As String, _
        Optional Module As VBComponent, _
        Optional TargetWorkbook As Workbook)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If DeclarationName = "" Then DeclarationName = CodepaneSelection
    If DeclarationName = "" Or InStr(1, DeclarationName, " ") > 0 Then Exit Sub
    Dim FilePath    As String
    FilePath = LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt"
    Dim code        As String
    On Error Resume Next
    code = TxtRead(FilePath)
    On Error GoTo 0

    If Len(code) = 0 Then    'CheckPath(filePath) = "I" Then
        On Error Resume Next
        code = TXTReadFromUrl(GITHUB_LIBRARY_DECLARATIONS & DeclarationName & ".txt")
        On Error GoTo 0
        If Len(code) > 0 And Not UCase(code) Like ("*NOT FOUND*") Then
            code = FormatVBA7(code)
            TxtOverwrite FilePath, code
        Else
            Debug.Print "File " & DeclarationName & ".txt not found localy or online"
            Exit Sub
        End If
    Else

    End If
    If InStr(1, WorkbookCode(TargetWorkbook), code, vbTextCompare) > 0 Then Exit Sub
    If Module Is Nothing Then Set Module = ModuleAddOrSet(TargetWorkbook, "vbArcImports", vbext_ct_StdModule)
    Module.CodeModule.AddFromString code

End Sub

Sub ImportProcedure( _
        Optional Procedure As String, _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If Procedure = "" Then Procedure = CodepaneSelection
    If Procedure = "" Or InStr(1, Procedure, " ") > 0 Then Exit Sub
    Dim ProcedurePath As String
    ProcedurePath = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"

    Dim code        As String
    On Error Resume Next
    code = TxtRead(ProcedurePath)
    On Error GoTo 0

    If Len(code) = 0 Then
        On Error Resume Next
        code = TXTReadFromUrl(GITHUB_LIBRARY_PROCEDURES & Procedure & ".txt")
        On Error GoTo 0
        If Len(code) > 0 And Not UCase(code) Like ("*NOT FOUND*") Then
            TxtOverwrite ProcedurePath, code
        Else
            Debug.Print "File " & Procedure & ".txt not found neither localy nor online"
            Exit Sub
        End If
    End If

    Dim filelastmod
    filelastmod = StringLastModified(code)
    Dim proclastmod

    If ProcedureExists(TargetWorkbook, Procedure) = True Then
        Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
        proclastmod = ProcedureLastModified(TargetWorkbook, Module, Procedure)
        If Overwrite = True Then
            If proclastmod = 0 Or proclastmod < filelastmod Then
                ProcedureReplace Module, Procedure, TxtRead(ProcedurePath)
            End If
        End If
    Else
        If Module Is Nothing Then
            '            Dim ModuleName As String
            '                ModuleName = StringProcedureAssignedModule(Code)
            '            If ModuleName = "" Then ModuleName = "vbArcImports"
            '            Set Module = ModuleAddOrSet(TargetWorkbook, ModuleName, vbext_ct_StdModule)
            Set Module = ModuleAddOrSet(TargetWorkbook, "vbArcImports", vbext_ct_StdModule)
        End If
        Module.CodeModule.AddFromFile ProcedurePath
    End If

    ImportProcedureDependencies Procedure, TargetWorkbook, Module, Overwrite
    '    ProcedureMoveToAssignedModule TargetWorkbook, Module, Procedure
End Sub

Sub ImportProcedureDependencies( _
        Optional Procedure As String, _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If Procedure = "" Then
        Dim cps     As String
        cps = CodepaneSelection
        If Len(cps) > 0 Then
            Procedure = cps
        Else
            Procedure = ActiveProcedure
        End If
        If Not ProcedureExists(TargetWorkbook, Procedure) Then Exit Sub
    End If
    On Error Resume Next
    If Module Is Nothing Then Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    If Module Is Nothing Then Exit Sub
    On Error GoTo 0
    Dim var
    Dim importfile  As String
    var = Split(ProcedureCode(TargetWorkbook, Module, Procedure), vbNewLine)
    var = Filter(var, "'@INCLUDE ")
    Dim TextLine    As Variant
    For Each TextLine In var
        TextLine = Trim(TextLine)
        If TextLine Like "'@INCLUDE *" Then
            importfile = Split(TextLine, " ")(2)
            importfile = Replace(importfile, vbNewLine, "")
            If TextLine Like "'@INCLUDE PROCEDURE *" Then
                ImportProcedure importfile, TargetWorkbook, Module, Overwrite
            ElseIf TextLine Like "'@INCLUDE CLASS *" Then
                ImportClass importfile, TargetWorkbook, Overwrite
            ElseIf TextLine Like "'@INCLUDE USERFORM *" Then
                ImportUserform importfile, TargetWorkbook, Overwrite
            ElseIf TextLine Like "'@INCLUDE DECLARATION *" Then
                ImportDeclaration importfile, Module, TargetWorkbook
            End If
        End If
    Next
End Sub

Sub ImportUserform( _
        Optional UserformName As String, _
        Optional TargetWorkbook As Workbook, _
        Optional Overwrite As Boolean)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If UserformName = "" Then UserformName = CodepaneSelection
    If UserformName = "" Or InStr(1, UserformName, " ") > 0 Then Exit Sub
    Dim FilePathFrM As String
    FilePathFrM = LOCAL_LIBRARY_USERFORMS & UserformName & ".frm"
    Dim FilePathFrX As String
    FilePathFrX = LOCAL_LIBRARY_USERFORMS & UserformName & ".frx"

    If CheckPath(FilePathFrM) = "I" Then
        On Error Resume Next
        Dim codeFrM As String
        codeFrM = TXTReadFromUrl(GITHUB_LIBRARY_USERFORMS & UserformName & ".frm")
        Dim codeFrX As String
        codeFrX = TXTReadFromUrl(GITHUB_LIBRARY_USERFORMS & UserformName & ".frx")
        On Error GoTo 0
        If Len(codeFrM) > 0 And Len(codeFrX) > 0 Then
            TxtOverwrite FilePathFrM, codeFrM
            TxtOverwrite FilePathFrX, codeFrX
        Else
            MsgBox "File " & UserformName & ".frm/.frx not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(UserformName, TargetWorkbook) Then
        If Overwrite = True Then
            TargetWorkbook.VBProject.VBComponents.Remove TargetWorkbook.VBProject.VBComponents(UserformName)
        Else
            Exit Sub
        End If
    End If
    TargetWorkbook.VBProject.VBComponents.Import FilePathFrM
End Sub

Public Function getLastCell(TargetWorksheet As Worksheet)
    Dim cell As Range
    Set cell = TargetWorksheet.Cells.Find(What:="*", _
                    after:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
    If cell Is Nothing Then Set cell = TargetWorksheet.Range("A1")
    Set getLastCell = cell
End Function

Public Function Len2( _
        ByVal val As Variant) _
        As Integer
    If IsArray(val) And Right(TypeName(val), 2) = "()" Then
        Len2 = UBound(val) - LBound(val) + 1
    ElseIf TypeName(val) = "String" Then
        Len2 = Len(val)
    ElseIf IsNumeric(val) Then
        Len2 = Len(CStr(val))
    Else
        Len2 = val.count
    End If
End Function

Function LinkedClasses( _
        TargetWorkbook As Workbook, _
        Module As VBComponent, _
        Procedure As String) As Collection

    Dim coll        As New Collection
    Dim var         As Variant
    var = classCallsOfModule(Module)
    Dim code        As String
    code = ProcedureCode(TargetWorkbook, Module, Procedure)
    Dim Keyword     As String
    Dim ClassName   As String
    Dim element     As Variant
    Dim i           As Long
    On Error Resume Next
    For i = LBound(var, 1) To UBound(var, 1)
        If InStr(1, code, var(i, 1)) > 0 Or InStr(1, code, var(i, 2)) > 0 Then
            coll.Add var(i, 1), var(i, 1)
        End If
    Next
    For Each element In ClassNames
        If InStr(1, code, element) > 0 Then
            coll.Add element, CStr(element)
        End If
    Next
    On Error GoTo 0
    Set LinkedClasses = coll
End Function

Function LinkedDeclarations( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As String) As Collection

    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop

    DeclarationsTableCreate TargetWorkbook

    Dim TargetWorksheet As Worksheet: Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    Dim coll        As New Collection
    Dim code        As String: code = ProcedureCode(TargetWorkbook, Module, Procedure)
    Dim element
    For Each element In DeclarationsTableKeywords
        If RegexTest(code, "\b ?" & CStr(element) & "\b") Then
            On Error Resume Next
            coll.Add CStr(element), CStr(element)
            On Error GoTo 0
        End If
    Next
    Set LinkedDeclarations = coll
End Function

Function LinkedProcedures( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional ProcedureName As String) As Collection
    If Not AssignCPSvariables(TargetWorkbook, Module, ProcedureName) Then Stop
    Dim Procedures  As Collection
    Set Procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim code        As String
    code = ProcedureCode(TargetWorkbook, Module, ProcedureName)
    Dim coll        As New Collection
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) <> UCase(CStr(ProcedureName)) Then
            If RegexTest(code, "\W" & CStr(Procedure) & "[.(\W]") = True Then
                coll.Add Procedure, CStr(Procedure)
            End If
        End If
    Next
    Set LinkedProcedures = coll
End Function

Function LinkedProceduresDeep( _
        ProcedureName As Variant, _
        TargetWorkbook As Workbook) As Collection

    Dim AllProcedures As Collection: Set AllProcedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Processed   As Collection: Set Processed = New Collection
    Dim CalledProcedures As Collection: Set CalledProcedures = New Collection

    Dim Procedure   As Variant
    Dim Module      As VBComponent

    Processed.Add CStr(ProcedureName), CStr(ProcedureName)
    On Error Resume Next
    For Each Procedure In LinkedProcedures(TargetWorkbook, , CStr(ProcedureName))
        CalledProcedures.Add CStr(Procedure), CStr(Procedure)
    Next
    On Error GoTo 0

    Dim CalledProceduresCount As Long
    CalledProceduresCount = CalledProcedures.count
    Dim element
repeat:
    For Each element In CalledProcedures
        If Not CollectionContains(Processed, , CStr(element)) Then
            On Error Resume Next
            For Each Procedure In LinkedProcedures(TargetWorkbook, , CStr(element))
                CalledProcedures.Add CStr(Procedure), CStr(Procedure)
            Next
            On Error GoTo 0
            Processed.Add CStr(element), CStr(element)
        End If
    Next
    If CalledProcedures.count > CalledProceduresCount Then
        CalledProceduresCount = CalledProcedures.count
        GoTo repeat
    End If

    Set LinkedProceduresDeep = CollectionSort(CalledProcedures)
End Function


Sub LinkedProceduresMoveHere(Optional Procedure As String)
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = ActiveCodepaneWorkbook
    If Not AssignProcedureVariable(TargetWorkbook, Procedure) Then Exit Sub
    Dim el
    For Each el In LinkedProceduresDeep(Procedure, TargetWorkbook)
        ProcedureMoveHere CStr(el)
    Next
End Sub




Function LinkedUserforms( _
        TargetWorkbook As Workbook, _
        Module As VBComponent, _
        Procedure As String) As Collection
    Dim coll        As New Collection
    Dim code        As String
    code = ProcedureCode(TargetWorkbook, Module, Procedure)
    Dim FormName
    For Each FormName In UserformNames(TargetWorkbook)
        If RegexTest(code, "\W" & FormName & "[.(\W]") = True Then coll.Add FormName    '& " " & "(Userform)"
    Next
    Set LinkedUserforms = coll
End Function

Function ModuleAddOrSet( _
        TargetWorkbook As Workbook, _
        TargetName As String, _
        moduleType As VBIDE.vbext_ComponentType) As VBComponent


    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module      As VBComponent
    On Error Resume Next
    Set Module = TargetWorkbook.VBProject.VBComponents(TargetName)
    On Error GoTo 0
    If Module Is Nothing Then
        Set Module = TargetWorkbook.VBProject.VBComponents.Add(moduleType)
        Module.Name = TargetName
    End If
    Set ModuleAddOrSet = Module
End Function




Function ModuleCode(Module As VBComponent) As String
    With Module.CodeModule
        If .countOfLines = 0 Then ModuleCode = "": Exit Function
        ModuleCode = .lines(1, .countOfLines)
    End With
End Function

Public Function ModuleExists( _
        TargetName As String, _
        TargetWorkbook As Workbook) As Boolean
    Dim Module      As VBComponent
    On Error Resume Next
    Set Module = TargetWorkbook.VBProject.VBComponents(TargetName)
    On Error GoTo 0
    ModuleExists = Not Module Is Nothing
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 25-10-2023 13:24    Alex                (Dependencies.bas > ModuleOfProcedure)

Public Function ModuleOfProcedure( _
        TargetWorkbook As Workbook, _
        ProcedureName As Variant) As VBComponent
'@LastModified 2310251324
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim lineNum As Long, NumProc As Long
    Dim Procedure   As String
    Dim Module      As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then
            With Module.CodeModule
                lineNum = .CountOfDeclarationLines + 1
                Do Until lineNum >= .countOfLines
                    Procedure = .ProcOfLine(lineNum, procKind)
                    If UCase(Procedure) = UCase(ProcedureName) Then
                        Set ModuleOfProcedure = Module
                        Exit Function
                    End If
                    lineNum = .procStartLine(Procedure, procKind) + .ProcCountLines(Procedure, procKind) + 1
                Loop
            End With
        End If
    Next Module
End Function

Function ModuleOrSheetName(Module As VBComponent) As String
    If Module.Type = vbext_ct_Document Then
        If Module.Name = "ThisWorkbook" Then
            ModuleOrSheetName = Module.Name
        Else
            ModuleOrSheetName = GetSheetByCodeName(WorkbookOfModule(Module), Module.Name).Name
        End If
    Else
        ModuleName = Module.Name
    End If
End Function

Function ModuleTypeToString(componentType As VBIDE.vbext_ComponentType) As String
    Select Case componentType
        Case vbext_ct_ActiveXDesigner
            ModuleTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ModuleTypeToString = "Class"
        Case vbext_ct_Document
            ModuleTypeToString = "Document"
        Case vbext_ct_MSForm
            ModuleTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ModuleTypeToString = "Module"
        Case Else
            ModuleTypeToString = "Unknown Type: " & CStr(componentType)
    End Select
End Function

Function ProcedureAssignedModule( _
        TargetWorkbook As Workbook, _
        Module As VBComponent, _
        Procedure As String) As VBComponent
    Dim ComponentName As Variant
    ComponentName = Split(ProcedureCode(TargetWorkbook, Module, Procedure), vbNewLine)
    ComponentName = Filter(ComponentName, "'@AssignedModule")
    If Len2(ComponentName) <> 1 Then Exit Function
    Dim ub          As Long
    ub = UBound(Split(ComponentName(0), " "))
    ComponentName = Split(ComponentName(0), " ")(ub)
    Set ProcedureAssignedModule = ModuleAddOrSet(TargetWorkbook, CStr(ComponentName), vbext_ct_StdModule)
End Function

Sub ProcedureAssignedModuleAdd( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    ProcedureLinesRemove "'@AssignedModule *", TargetWorkbook, Module, Procedure
    Module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(Module, Procedure), _
            "'@AssignedModule " & Module.Name
End Sub

Function ProcedureBodyLineFirst( _
        Module As VBComponent, _
        Procedure As String) As Long
    ProcedureBodyLineFirst = ProcedureTitleLineFirst(Module, Procedure) + ProcedureTitleLineCount(Module, Procedure)
End Function

Function ProcedureBodyLineFirstAfterComments( _
        Module As VBComponent, _
        Procedure As String) As Long
    Dim N           As Long
    Dim s           As String
    For N = ProcedureBodyLineFirst(Module, Procedure) To Module.CodeModule.countOfLines
        s = Trim(Module.CodeModule.lines(N, 1))
        If s = vbNullString Then
            Exit For
        ElseIf Left(s, 1) = "'" Then
        ElseIf Left(s, 3) = "Rem" Then
        ElseIf Right(Trim(Module.CodeModule.lines(N - 1, 1)), 1) = "_" Then
        ElseIf Right(s, 1) = "_" Then
        Else
            Exit For
        End If
    Next N
    ProcedureBodyLineFirstAfterComments = N
End Function



Public Function ProcedureCode( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As Variant, _
        Optional IncludeHeader As Boolean = True) As String
    If Not AssignCPSvariables(TargetWorkbook, Module, CStr(Procedure)) Then Exit Function
    Dim lProcStart  As Long
    Dim lProcBodyStart As Long
    Dim lProcNoLines As Long
    Const vbext_pk_Proc = 0
    On Error GoTo Error_Handler
    lProcStart = Module.CodeModule.procStartLine(Procedure, vbext_pk_Proc)
    lProcBodyStart = Module.CodeModule.ProcBodyLine(Procedure, vbext_pk_Proc)
    lProcNoLines = Module.CodeModule.ProcCountLines(Procedure, vbext_pk_Proc)
    If IncludeHeader = True Then
        ProcedureCode = Module.CodeModule.lines(lProcStart, lProcNoLines)
    Else
        lProcNoLines = lProcNoLines - (lProcBodyStart - lProcStart)
        ProcedureCode = Module.CodeModule.lines(lProcBodyStart, lProcNoLines)
    End If
'Error_Handler_Exit:
'    On Error Resume Next
'    Exit Function
Error_Handler:
'possible cause the proc exists but is single line
'    Debug.Print "Error Source: ProcedureCode" & vbCrLf & _
'            "Error Description: " & Err.Description & _
'            Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Line No: " & Erl)
'    Resume Error_Handler_Exit
End Function

Function ProcedureExists( _
        TargetWorkbook As Workbook, _
        ProcedureName As Variant) As Boolean
    Dim Procedures  As Collection
    Set Procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) = UCase(ProcedureName) Then
            ProcedureExists = True
            Exit Function
        End If
    Next
End Function

Function ProcedureLastModAdd( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As String, _
        Optional ModificationDate As Double)



    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Exit Function
    If ModificationDate = 0 Then ModificationDate = Format(Now, "yymmddhhnn")
    Dim LastModLine As Long
    LastModLine = ProcedureLineContaining(Module, Procedure, "'@LastModified *")
    If LastModLine = 0 Then GoTo PASS
    Dim LDate       As Double
    LDate = Split(Module.CodeModule.lines(LastModLine, 1), " ")(1)
    ProcedureLinesRemove "'@LastModified *", TargetWorkbook, Module, Procedure
PASS:
    Module.CodeModule.InsertLines ProcedureBodyLineFirst(Module, Procedure), _
            "'@LastModified " & ModificationDate

    ProcedureLastModAdd = ModificationDate
End Function

Function ProcedureLastModified( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    ProcedureLastModified = StringLastModified(ProcedureCode(TargetWorkbook, Module, Procedure))
End Function

Function ProcedureLinesCount( _
        Module As VBComponent, _
        Procedure As String) As Long
    ProcedureLinesCount = Module.CodeModule.ProcCountLines(Procedure, vbext_pk_Proc)
End Function

Public Function ProcedureLinesFirst( _
        Module As VBComponent, _
        Procedure As String) As Long
    Dim procKind    As VBIDE.vbext_ProcKind
    procKind = vbext_pk_Proc
    ProcedureLinesFirst = Module.CodeModule.procStartLine(Procedure, procKind)
End Function


Public Function ProcedureLinesLast( _
        Module As VBComponent, _
        Procedure As String, _
        Optional IncludeTail As Boolean) As Long
    Dim procKind    As VBIDE.vbext_ProcKind
    procKind = vbext_pk_Proc
    Dim startAt     As Long
    startAt = Module.CodeModule.procStartLine(Procedure, procKind)
    Dim CountOf     As Long
    CountOf = Module.CodeModule.ProcCountLines(Procedure, procKind)
    Dim endAt       As Long
    endAt = startAt + CountOf - 1
    If Not IncludeTail Then
        Do While Not Trim(Module.CodeModule.lines(endAt, 1)) Like "End *"
            endAt = endAt - 1
        Loop
    End If
    ProcedureLinesLast = endAt
End Function

Sub ProcedureLinesRemove( _
        myCriteria As String, _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop

    Dim code        As String
    Dim i           As Long
    For i = ProcedureLinesLast(Module, Procedure) To ProcedureLinesFirst(Module, Procedure) Step -1
        code = Trim(Module.CodeModule.lines(i, 1))
        If code Like myCriteria Then Module.CodeModule.DeleteLines i
    Next
End Sub

Sub ProcedureLinesRemoveInclude( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    ProcedureLinesRemove "'@INCLUDE", TargetWorkbook, Module, Procedure
End Sub


Sub ProcedureMoveHere( _
        Optional Procedure As String)


    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = ActiveCodepaneWorkbook
    If Not AssignProcedureVariable(TargetWorkbook, Procedure) Then Exit Sub
    Dim Module      As VBComponent
    Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    Dim s           As String
    s = ProcedureCode(TargetWorkbook, Module, Procedure)

    If InStr(1, s, "'@AssignedModule") = 0 Then
        ProcedureAssignedModuleAdd TargetWorkbook, Module, Procedure
        s = ProcedureCode(TargetWorkbook, Module, Procedure)
    End If

    Dim sl As Long, cl As Long
    sl = ProcedureLinesFirst(Module, Procedure)
    cl = ProcedureLinesLast(Module, Procedure, False) - sl + 1
    ActiveModule.CodeModule.InsertLines ProcedureLinesLast(Module, ActiveProcedure, True) + 1, s
    Module.CodeModule.DeleteLines sl, cl
End Sub

Sub ProcedureMoveToAssignedModule( _
        Optional TargetWorkbook As Workbook, _
        Optional Module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Exit Sub
    Dim MoveToModule As VBComponent
    Set MoveToModule = ProcedureAssignedModule(TargetWorkbook, Module, Procedure)
    If MoveToModule Is Nothing Then Exit Sub
    ProcedureMoveToModule TargetWorkbook, Module, Procedure, MoveToModule
End Sub

Sub ProcedureMoveToModule( _
        TargetWorkbook As Workbook, _
        Module As VBComponent, _
        Procedure As String, _
        MoveToModule As VBComponent)
    Dim code        As String
    code = ProcedureCode(TargetWorkbook, Module, Procedure)
    Dim StartLine   As Long
    StartLine = ProcedureLinesFirst(Module, Procedure)
    Dim CountLines  As Long
    CountLines = ProcedureLinesCount(Module, Procedure)
    MoveToModule.CodeModule.InsertLines MoveToModule.CodeModule.countOfLines + 1, vbNewLine & code
    Module.CodeModule.DeleteLines StartLine, CountLines

End Sub

Public Sub ProcedureReplace( _
        Module As VBComponent, _
        Procedure As String, _
        code As String)

    Dim StartLine   As Integer
    Dim NumLines    As Integer
    With Module.CodeModule
        StartLine = .procStartLine(Procedure, vbext_pk_Proc)
        NumLines = .ProcCountLines(Procedure, vbext_pk_Proc)
        .DeleteLines StartLine, NumLines
        .InsertLines StartLine, code
    End With
End Sub

Function ProcedureTitle( _
        Module As VBComponent, _
        Procedure As String) As String
    Dim titleLine   As Long
    titleLine = ProcedureTitleLineFirst(Module, Procedure)
    Dim title       As String
    title = Module.CodeModule.lines(titleLine, 1)
    Dim counter     As Long
    counter = 1
    Do While Right(title, 1) = "_"
        counter = counter + 1
        title = Module.CodeModule.lines(titleLine, counter)
    Loop

    ProcedureTitle = title
End Function

Function ProcedureTitleLineCount( _
        Module As VBComponent, _
        Procedure As String) As Long

    ProcedureTitleLineCount = ProcedureTitleLineLast(Module, Procedure) - ProcedureTitleLineFirst(Module, Procedure) + 1
End Function



Public Function ProcedureTitleLineFirst( _
        Module As VBComponent, _
        Procedure As String) As Long
    ProcedureTitleLineFirst = Module.CodeModule.ProcBodyLine(Procedure, vbext_pk_Proc)
End Function

Function ProcedureTitleLineLast( _
        Module As VBComponent, _
        Procedure As String) As Long
    ProcedureTitleLineLast = ProcedureTitleLineFirst(Module, Procedure) + UBound(Split(ProcedureTitle(Module, Procedure), vbNewLine))
End Function

Public Function ProceduresOfModule( _
        Module As VBComponent) As Collection
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim Procedure   As String
    With Module.CodeModule
        lineNum = .CountOfDeclarationLines + 1
        Do Until lineNum >= .countOfLines
            ProcedureAs = .ProcOfLine(lineNum, procKind)
            coll.Add ProcedureAs
            lineNum = .procStartLine(ProcedureAs, procKind) + .ProcCountLines(ProcedureAs, procKind) + 1
        Loop
    End With
    Set ProceduresOfModule = coll
End Function

Function ProceduresOfTXT( _
        code As String) As Collection


    code = Replace(code, vbNewLine, vbLf)
    Dim var
    var = Split(code, vbLf)

    Dim out
    out = ArrayAppend(Filter(var, "Sub" & Space(1), True, vbBinaryCompare), Filter(var, "Function ", True, vbBinaryCompare))
    If TypeName(out) = "Empty" Then Exit Function
    out = Filter(out, "(", True)
    out = Filter(out, "Declare", False)
    out = Filter(out, Chr(34) & "Sub", False)
    out = Filter(out, Chr(34) & "Function", False)
    out = Filter(out, "End Sub", False)
    out = Filter(out, "End Function", False)

    Dim i           As Long
    For i = LBound(out) To UBound(out)
        out(i) = Left(out(i), InStr(1, out(i), "(") - 1)
        out(i) = Replace(out(i), "Private ", "")
        out(i) = Replace(out(i), "Public ", "")
        out(i) = Replace(out(i), "Sub ", "")
        out(i) = Replace(out(i), "Function ", "")
        If UBound(Split(out(i), " ")) > 0 Then
            out(i) = ""
        End If
    Next

    ArrayQuickSort out
    out = cleanArray(out)
    out = ArrayDuplicatesRemove(out)
    Set ProceduresOfTXT = ArrayToCollection(out)
End Function

Sub CallSeparateProcedures()
    Dim FilePath    As Variant
    FilePath = DataFilePicker("*.txt", False)
    If FilePath = vbNullString Then Exit Sub
    Dim OutputFolder As Variant
    OutputFolder = SelectFolder(Left(FilePath, InStrRev(FilePath, "\")))

    TxtSeparateProcedures FilePath, OutputFolder

End Sub

Sub TxtSeparateProcedures(FilePath As Variant, Optional OutputFolder As Variant)

    '@AssignedModule F_FileFolder
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE PROCEDURE TxtRead
    Dim FName       As String
    If OutputFolder = "" Then
        OutputFolder = Left(FilePath, InStrRev(FilePath, "\"))
    Else
        FoldersCreate CStr(OutputFolder)
    End If
    Dim code        As Variant
    code = Split(TxtRead(FilePath), vbLf)
    Dim out         As String
    Dim i           As Long
    For i = LBound(code) To UBound(code)

        out = IIf(out = "", code(i), out & code(i)) & vbNewLine
        If RegexTest(code(i), "Sub ") _
                And Not code(i) Like Chr(34) & "*Sub*" Then
            FName = Split(code(i), "Sub ")(1)
            FName = Trim(Split(FName, "(")(0)) & ".txt"
        ElseIf RegexTest(code(i), "Function ") _
                And Not code(i) Like Chr(34) & "*Function*" Then
            FName = Split(code(i), "Function ")(1)
            FName = Trim(Split(FName, "(")(0)) & ".txt"
        End If
        If Trim(code(i)) = "End Sub" Or Trim(code(i)) = "End Function" Then
            TxtOverwrite OutputFolder & FName, out
            out = ""
            FName = ""
        End If
    Next

End Sub

Function ProceduresOfWorkbook( _
        TargetWorkbook As Workbook, _
        Optional ExcludeDocument As Boolean = True, _
        Optional ExcludeClass As Boolean = True, _
        Optional ExcludeForm As Boolean = True) As Collection
    Dim Module      As VBComponent
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim ProcedureName As String
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If ExcludeClass = True And Module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If ExcludeDocument = True And Module.Type = vbext_ct_Document Then GoTo SKIP
        If ExcludeForm = True And Module.Type = vbext_ct_MSForm Then GoTo SKIP
        With Module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .countOfLines
                ProcedureName = .ProcOfLine(lineNum, procKind)
                If InStr(1, ProcedureName, "_") = 0 Then
                    coll.Add ProcedureName
                End If
                lineNum = .procStartLine(ProcedureName, procKind) + .ProcCountLines(ProcedureName, procKind) + 1
            Loop
        End With
SKIP:
    Next Module
    Set ProceduresOfWorkbook = coll
End Function

Sub ProjetFoldersCreate()
    Dim element
    For Each element In vbarcFolders
        FoldersCreate CStr(element)
    Next
End Sub

Public Function RegexTest( _
        ByVal string1 As String, _
        ByVal stringPattern As String, _
        Optional ByVal globalFlag As Boolean, _
        Optional ByVal ignoreCaseFlag As Boolean, _
        Optional ByVal multilineFlag As Boolean) As Boolean
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = globalFlag
        .IgnoreCase = ignoreCaseFlag
        .MultiLine = multilineFlag
        .pattern = stringPattern
    End With
    RegexTest = regex.test(string1)
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 11:03    Alex                (Dependencies.bas > StringLastModified)

Function StringLastModified(txt As String)
'@LastModified 2308221103

    Dim code        As Variant
    code = Filter(Filter(Split(txt, vbLf), "'@LastModified ", True), """", False)
    If ArrayAllocated(code) Then
        Dim lastDate As Variant
        If Trim(code(0)) Like "'@LastModified *" Then
            lastDate = Split(code(0), " ")(1)
            lastDate = DateSerial(Left(lastDate, 2), Mid(lastDate, 3, 2), Mid(lastDate, 5, 2)) _
                    & " " & TimeSerial(Mid(lastDate, 7, 2), Mid(lastDate, 9, 2), 0)
            StringLastModified = Split(code(0), " ")(1)
        End If
    Else

    End If
End Function

Function StringProcedureAssignedModule(txt As String) As String
    Dim ComponentName As Variant
    ComponentName = Split(txt, vbLf)
    ComponentName = Filter(ComponentName, "'@AssignedModule")
    If Not ArrayAllocated(ComponentName) Then Exit Function
    Dim ub          As Long
    ub = UBound(Split(ComponentName(0), " "))
    ComponentName = Split(ComponentName(0), " ")(ub)
    StringProcedureAssignedModule = ComponentName
End Function



Function TXTReadFromUrl(url As String) As String
    On Error GoTo Err_GetFromWebpage
    Dim objWeb      As Object
    Dim strXML      As String
    Set objWeb = CreateObject("Msxml2.ServerXMLHTTP")
    objWeb.Open "GET", url, False
    objWeb.setRequestHeader "Content-Type", "text/xml"
    objWeb.setRequestHeader "Cache-Control", "no-cache"
    objWeb.setRequestHeader "Pragma", "no-cache"
    objWeb.send
    Do While objWeb.readyState <> 4
        DoEvents
    Loop
    strXML = objWeb.responseText
    TXTReadFromUrl = strXML
End_GetFromWebpage:
    Set objWeb = Nothing
    Exit Function
Err_GetFromWebpage:
    MsgBox Err.Description & " (" & Err.Number & ")"
    Resume End_GetFromWebpage
End Function

Sub TxtOverwrite(sFile As String, sText As String)
    On Error GoTo ERR_HANDLER
    Dim FileNumber  As Integer
    FileNumber = FreeFile
    Open sFile For Output As #FileNumber
    Print #FileNumber, sText
    Close #FileNumber
Exit_Err_Handler:
    Exit Sub
ERR_HANDLER:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: TxtOverwrite" & vbCrLf & _
            "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Sub

Sub TxtPrepend(FilePath As String, txt As String)
    Dim s           As String
    s = TxtRead(FilePath)
    TxtOverwrite FilePath, txt & vbNewLine & s
End Sub


Sub CallTxtPrependContainedProcedures()
    Dim FilePath    As Variant
    FilePath = DataFilePicker("*.txt", False)
    If FilePath = vbNullString Then Exit Sub
    TxtPrependContainedProcedures CStr(FilePath)
End Sub

Sub TxtPrependContainedProcedures(fileName As String)
    Dim s           As String: s = TxtRead(fileName)
    Dim v           As New Collection
    Set v = ProceduresOfTXT(s)
    If v.count = 0 Then Exit Sub
    Dim Line        As String: Line = String(30, "'")
    TxtPrepend fileName, _
            "'Contains the following " & "#" & v.count & " procedures " & vbNewLine & Line & vbNewLine & _
            "'" & collectionToString(v, vbNewLine & "'") & vbNewLine & Line & vbNewLine & vbNewLine
End Sub

Function TxtRead(sPath As Variant) As String
    Dim sTXT        As String
    If Dir(sPath) = "" Then
        Debug.Print "File was not found."
        Debug.Print sPath
        Exit Function
    End If
    Open sPath For Input As #1
    Do Until EOF(1)
        Line Input #1, sTXT
        TxtRead = TxtRead & sTXT & vbLf
    Loop
    Close
    If Len(TxtRead) = 0 Then
        TxtRead = ""
    Else
        TxtRead = Left(TxtRead, Len(TxtRead) - 1)
    End If
End Function

Function URLExists(url) As Boolean
    Dim Request     As Object
    Dim FF          As Integer
    Dim rc          As Variant

    On Error GoTo EndNow
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")

    With Request
        .Open "GET", url, False
        .send
        rc = .statusText
    End With
    Set Request = Nothing
    If rc = "OK" Then URLExists = True

    Exit Function
EndNow:
End Function

Function UserformNames(TargetWorkbook As Workbook)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Set UserformNames = ComponentNames(vbext_ct_MSForm, TargetWorkbook)
End Function






Function WorkbookCode(TargetWorkbook) As String
    If TypeName(TargetWorkbook) <> "Workbook" Then Stop
    Dim Module      As VBComponent
    Dim txt
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.CodeModule.countOfLines > 0 Then
            txt = txt & _
                    vbNewLine & _
                    "'" & String(10, "=") & ModuleOrSheetName(Module) & " (" & Module.Type & ") " & String(10, "=") & _
                    vbNewLine & _
                    ModuleCode(Module)
        End If
    Next
    WorkbookCode = txt
End Function


Function WorkbookOfModule(vbComp As VBComponent) As Workbook
    Set WorkbookOfModule = WorkbookOfProject(vbComp.Collection.Parent)
End Function

Function WorkbookOfProject(vbProj As VBProject) As Workbook
    tmpStr = vbProj.fileName
    tmpStr = Right(tmpStr, Len(tmpStr) - InStrRev(tmpStr, "\"))
    Set WorkbookOfProject = Workbooks(tmpStr)
End Function



Function WorksheetExists(SheetName As String, TargetWorkbook As Workbook) As Boolean
    Dim TargetWorksheet As Worksheet
    On Error Resume Next
    Set TargetWorksheet = TargetWorkbook.Sheets(SheetName)
    On Error GoTo 0
    WorksheetExists = Not TargetWorksheet Is Nothing
End Function

Function classCallsOfModule(Module As VBComponent) As Variant


    Dim code        As Variant
    Dim element     As Variant
    Dim Keyword     As Variant
    Dim var         As Variant
    ReDim var(1 To 2, 1 To 1)
    Dim counter     As Long
    counter = 0
    If Module.CodeModule.CountOfDeclarationLines > 0 Then
        code = Module.CodeModule.lines(1, Module.CodeModule.CountOfDeclarationLines)
        code = Replace(code, "_" & vbNewLine, "")
        code = Split(code, vbNewLine)
        code = Filter(code, " As ", , vbTextCompare)
        For Each element In code
            element = Trim(element)
            If element Like "* As *" Then
                Keyword = Split(element, " As ")(0)
                Keyword = Split(Keyword, " ")(UBound(Split(Keyword, " ")))
                element = Split(element, " As ")(1)
                element = Replace(element, "New ", "")

                For Each ClassName In ClassNames
                    If element = ClassName Then

                        ReDim Preserve var(1 To 2, 1 To counter + 1)
                        var(1, UBound(var, 2)) = element
                        var(2, UBound(var, 2)) = Keyword
                        counter = counter + 1
                    End If
                Next
            End If
        Next
        If var(1, 1) <> "" Then

            If UBound(var, 2) > 1 Then
                classCallsOfModule = WorksheetFunction.Transpose(var)
            Else
                Dim VAR2(1 To 1, 1 To 2)
                VAR2(1, 1) = var(1, 1)
                VAR2(1, 2) = var(2, 1)
                classCallsOfModule = VAR2
            End If
        End If
    End If

End Function

Function collectionToString(coll As Collection, delim As String) As String
    Dim element
    Dim out         As String
    For Each element In coll
        out = IIf(out = "", element, out & delim & element)
    Next
    collectionToString = out
End Function

Function getDeclarations( _
        wb As Workbook, _
        Optional includeScope As Boolean, _
        Optional includeType As Boolean, _
        Optional includeKeywords As Boolean, _
        Optional includeDeclarations As Boolean, _
        Optional includeComponentName As Boolean, _
        Optional includeComponentType As Boolean) As Collection

    Dim ComponentCollection As New Collection
    Dim ComponentTypecollection As New Collection
    Dim DeclarationsCollection As New Collection
    Dim KeywordsCollection As New Collection
    Dim Output      As New Collection
    Dim ScopeCollection As New Collection
    Dim TypeCollection As New Collection

    Dim element     As Variant
    Dim OriginalDeclarations As Variant
    Dim str         As Variant

    Dim tmp         As String
    Dim helper      As String
    Dim i           As Long

    Dim Module      As VBComponent
    For Each Module In wb.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule Or Module.Type = vbext_ct_MSForm Then
            If Module.CodeModule.CountOfDeclarationLines > 0 Then
                str = Module.CodeModule.lines(1, Module.CodeModule.CountOfDeclarationLines)
                str = Replace(str, "_" & vbNewLine, "")
                OriginalDeclarations = str
                tmp = str
                Do While InStr(1, str, "End Type") > 0
                    tmp = Mid(str, InStr(1, str, "Type "), InStr(1, str, "End Type") - InStr(1, str, "Type ") + 8)
                    str = Replace(str, tmp, Split(tmp, vbNewLine)(0))
                Loop
                Do While InStr(1, str, "End Enum") > 0
                    tmp = Mid(str, InStr(1, str, "Enum "), InStr(1, str, "End Enum") - InStr(1, str, "Enum ") + 8)
                    str = Replace(str, tmp, Split(tmp, vbNewLine)(0))
                Loop
                Do While InStr(1, str, "  ") > 0
                    str = Replace(str, "  ", " ")
                Loop

                str = Split(str, vbNewLine)
                tmp = OriginalDeclarations

                For Each element In str
                    If Len(CStr(element)) > 0 And Not Trim(CStr(element)) Like "'*" And Not Trim(CStr(element)) Like "Rem*" Then
                        If RegexTest(CStr(element), "\b ?Enum \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Enum")
                            DeclarationsCollection.Add DeclarationsKeywordSubstring(tmp, , "Enum " & KeywordsCollection.Item(KeywordsCollection.count), "End Enum", , , True)
                            TypeCollection.Add "Enum"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add ModuleTypeToString(Module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Type \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Type")
                            DeclarationsCollection.Add DeclarationsKeywordSubstring(tmp, , "Type " & KeywordsCollection.Item(KeywordsCollection.count), "End Type", , , True)
                            TypeCollection.Add "Type"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add ModuleTypeToString(Module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf InStr(1, CStr(element), "Const ", vbTextCompare) > 0 Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Const")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Const"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add ModuleTypeToString(Module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Sub \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Sub")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Sub"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add ModuleTypeToString(Module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Function \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Function")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Function"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add ModuleTypeToString(Module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf element Like "*(*) As *" Then
                            helper = Left(element, InStr(1, CStr(element), "(") - 1)
                            helper = Mid(helper, InStrRev(helper, " ") + 1)
                            KeywordsCollection.Add helper
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Other"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add ModuleTypeToString(Module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf element Like "* As *" Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", , "As")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Other"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add ModuleTypeToString(Module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        Else
                        End If
                    End If
                Next element
            End If
        End If
    Next Module

    If includeScope = True Then Output.Add ScopeCollection
    If includeType = True Then Output.Add TypeCollection
    If includeKeywords = True Then Output.Add KeywordsCollection
    If includeDeclarations = True Then Output.Add DeclarationsCollection
    If includeComponentType = True Then Output.Add ComponentTypecollection
    If includeComponentName = True Then Output.Add ComponentCollection

    Set getDeclarations = Output
End Function

Function getLastRow(TargetSheet As Worksheet)
    Dim LastCell    As Range
    Set LastCell = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    getLastRow = LastCell.row
End Function

Function vbarcFolders() As Collection
    Dim coll        As New Collection
    coll.Add LOCAL_LIBRARY_PROCEDURES
    coll.Add LOCAL_LIBRARY_CLASSES
    coll.Add LOCAL_LIBRARY_USERFORMS
    coll.Add LOCAL_LIBRARY_DECLARATIONS
    Set vbarcFolders = coll
End Function

Function ProcedureLineContaining(Module As VBComponent, Procedure As String, this As String) As Long
    Dim i           As Long
    For i = ProcedureLinesFirst(Module, Procedure) To ProcedureLinesLast(Module, Procedure)
        If Module.CodeModule.lines(i, 1) Like this Then
            ProcedureLineContaining = i
            Exit Function
        End If
    Next
End Function

Public Sub ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Sub

Public Function CLIP(Optional StoreText As String) As String
    Dim x           As Variant
    x = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText)
                    .SetData "text", x
                Case Else
                    CLIP = .GetData("text")
            End Select
        End With
    End With
End Function


