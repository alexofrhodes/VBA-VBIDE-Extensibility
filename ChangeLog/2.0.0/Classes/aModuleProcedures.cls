VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aModuleProcedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize         As aModuleProcedures  :
' Fun Items              As Variant            :
' Fun Names              As Collection         :
' Fun PrivateProcedures  As Collection         :
' Fun PublicProcedures   As Collection         :
' Sub Export                                   :
' Sub List                                     :
' Sub ListPublic                               :
' Sub SortAZ                                   :
' Sub SortByKind                               :
' Sub SortByScope                              :
' Sub Update                                   :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit


Private oModule As aModule

Public Function Initialize(Module As aModule) As aModuleProcedures
    Set oModule = Module
    Set Initialize = Me
End Function

Public Function Items(Optional index As Variant)
    If IsMissing(index) Then
        Set Items = Procedures
    ElseIf IsNumeric(index) Then
        Set Items = Procedures(index)
    Else
        Dim ap As aProcedure
        For Each ap In Procedures
            If ap.Name = index Then Set Items = ap: Exit For
        Next
    End If
End Function

Private Function Procedures() As Collection
    Dim lineNum     As Long
    Dim ProcedureName As String
    Dim procKind    As Long
    Dim coll        As New Collection
    Dim i           As Long
    Dim cProcedure  As aProcedure
    With oModule.Component.CodeModule

        lineNum = .CountOfDeclarationLines + 1
        Do Until lineNum >= .countOfLines
            On Error Resume Next
            For i = 0 To 4
                procKind = i
                ProcedureName = .ProcOfLine(lineNum, procKind)
                If ProcedureName <> vbNullString Then Exit For
            Next
            On Error GoTo 0
            'Err.clear
            If ProcedureName = vbNullString Then Exit Function

            'Possible Error:
            '---------------
            ' Procedures with an underscore _ which is used in events
            ' and Private Procedures
            ' may have the same name in different components

            If InStr(1, ProcedureName, "_") = 0 Then    '>0 and ExcludeUnderscore then goto SKIP

                Set cProcedure = New aProcedure
                cProcedure.Initialize oModule.WorkbookObject, oModule.Component, ProcedureName, procKind
                On Error Resume Next
                coll.Add cProcedure    ', cProcedure.Name
                '                    coll.Add ProcedureName, ProcedureName
                '                    If Err.Number > 0 Then Debug.Print ProcedureName & " exists multiple times, added only once, may lead to error"
                '                    Err.clear
                On Error GoTo 0
                lineNum = .procStartLine(ProcedureName, procKind) + .ProcCountLines(ProcedureName, procKind) + 1
            End If
            lineNum = .procStartLine(ProcedureName, procKind) + .ProcCountLines(ProcedureName, procKind) + 1
            ProcedureName = vbNullString
        Loop
    End With
    Set Procedures = coll
End Function

Function Names() As Collection
    Dim ap          As aProcedure
    Dim coll        As New Collection
    Dim ProcedureName As String
    For Each ap In Procedures
        ProcedureName = ap.Name
        On Error Resume Next
        coll.Add ProcedureName, ProcedureName
        On Error GoTo 0
        If Err.Number > 0 Then Debug.Print ProcedureName & " exists multiple times, added only once, may lead to error"
        Err.clear
    Next
    Set Names = coll
End Function

Function PublicProcedures() As Collection
    Dim ap          As aProcedure
    Dim coll        As New Collection
    Dim ProcedureName As String
    For Each ap In Procedures
        If ap.Scope.ToString = "Public" Then coll.Add ap
    Next
    Set PublicProcedures = coll
End Function

Function PrivateProcedures() As Collection
    Dim ap          As aProcedure
    Dim coll        As New Collection
    Dim ProcedureName As String
    For Each ap In Procedures
        If ap.Scope.ToString = "Private" Then coll.Add ap
    Next
    Set PrivateProcedures = coll
End Function


Public Sub List()
    Dim coll        As New Collection
    Set coll = Names
    If coll.count = 0 Then Exit Sub

    Dim txt         As String

    Dim tmp
    Dim arr:    arr = aCollection.Init(coll).ToArray
    Dim ll:     ll = LargestLength(coll)
    Dim ap As aProcedure
    For Each ap In Procedures
        Dim ks As String
        ks = ap.KindAsString
        tmp = "'" & Space(4) & Left(ks, 3) & Space(1) & ap.Name
        Dim rt As String: rt = ap.returnType: If rt <> "" Then tmp = tmp & " As " & ap.returnType
        
        tmp = tmp & " :"

        Dim Desc As String, index As Long
        index = ap.lines.LikeThis("*'@Description*", ap.lines.FirstOfBody)
        If index > 0 Then
            Desc = ap.Parent.CodeModule.lines(index, 1)
            tmp = tmp & " " & Desc
        End If
        txt = txt & tmp & vbNewLine
    Next
    arr = Split(txt, vbNewLine)
    arr = cleanArray(arr)
    txt = Join(arr, vbNewLine)
    txt = StringFormatAlignRowsElements(txt, " As ", True)
    txt = StringFormatAlignRowsElements(txt, ":", True)
    oModule.Folders.Overwrite "ListOfProcedures", txt
End Sub

Public Sub ListPublic()

    Dim coll        As New Collection
    Set coll = PublicProcedures
    If coll.count = 0 Then Exit Sub

    Dim txt         As String

    Dim tmp
    Dim arr
    Dim ap As aProcedure
    Dim nameCollection As New Collection
    For Each ap In coll
        On Error Resume Next
        nameCollection.Add ap.Name, ap.Name
        On Error GoTo 0
    Next
    Dim ll:     ll = LargestLength(nameCollection)
    For Each ap In Procedures
        If ap.Scope.ToString = "Public" Then
            Dim ks As String
            ks = ap.KindAsString
            tmp = "'" & Space(4) & Left(ks, 3) & Space(1) & ap.Name
            Dim rt As String: rt = ap.returnType: If rt <> "" Then tmp = tmp & " As " & ap.returnType
            tmp = tmp & " :"
            Dim Desc As String, index As Long
            index = ap.lines.LikeThis("*'@Description*", ap.lines.FirstOfBody)
            If index > 0 Then
                If Trim(ap.oModule.CodeModule.lines(index, 1)) Like "'@Description*" Then
                    Desc = ap.Parent.CodeModule.lines(index, 1)
                    tmp = tmp & " " & Desc
                End If
            End If
            txt = txt & tmp & vbNewLine
        End If
    Next
    arr = Split(txt, vbNewLine)
    arr = cleanArray(arr)
    ArraySort arr
    txt = Join(arr, vbNewLine)
    txt = StringFormatAlignRowsElements(txt, " As ", True)
    txt = StringFormatAlignRowsElements(txt, ":", True)

    oModule.Folders.Overwrite "ListOfProcedures", txt
End Sub


Public Sub SortByKind()
    If oModule.Component.CodeModule.countOfLines = 0 Then Exit Sub
    Dim coll        As New Collection: Set coll = Procedures
    Dim var         As Variant: ReDim var(1 To coll.count)
    Dim p           As aProcedure
    Dim StartLine   As Long
    Dim i           As Long
    For i = 1 To coll.count
        Set p = coll(i)
        var(i) = p.Name & "-" & p.KindAsLong
        If i = 1 Then StartLine = p.lines.first
    Next
    ArrayQuickSort var
    Dim h           As String: h = oModule.Header
    Dim S           As String
    Dim TheSubs As String, TheFunctions As String
    For i = LBound(var) To UBound(var)
        Set p = aProcedure.Initialize(oModule.WorkbookObject, oModule.Component, CStr(Split(var(i), "-")(0)), CLng(Split(var(i), "-")(1)))
        S = p.code.All
        If p.KindAsString = "Sub" Then
            TheSubs = TheSubs & IIf(TheSubs <> "", vbNewLine, "") & S
        Else
            TheFunctions = TheFunctions & IIf(TheFunctions <> "", vbNewLine, "") & S
        End If
    Next i
    oModule.CodeRemove
    oModule.Component.CodeModule.AddFromString h & vbLf & TheSubs & vbLf & TheFunctions

End Sub

Public Sub SortByScope()
    If oModule.Component.CodeModule.countOfLines = 0 Then Exit Sub
    Dim coll        As New Collection: Set coll = Procedures
    Dim var         As Variant: ReDim var(1 To coll.count)
    Dim p           As aProcedure
    Dim StartLine   As Long
    Dim i           As Long
    For i = 1 To coll.count
        Set p = coll(i)
        var(i) = p.Name & "-" & p.KindAsLong
        If i = 1 Then StartLine = p.lines.first
    Next
    ArrayQuickSort var
    Dim h           As String: h = oModule.Header
    Dim S           As String
    Dim ThePublic As String, ThePrivate As String
    For i = LBound(var) To UBound(var)
        Set p = aProcedure.Initialize(oModule.WorkbookObject, oModule.Component, CStr(Split(var(i), "-")(0)), CLng(Split(var(i), "-")(1)))
        S = p.code.All
        If p.Scope.ToString = "Public" Then
            ThePublic = ThePublic & IIf(ThePublic <> "", vbNewLine, "") & S
        Else
            ThePrivate = ThePrivate & IIf(ThePublic <> "", vbNewLine, "") & S
        End If
    Next i
    oModule.CodeRemove
    oModule.Component.CodeModule.AddFromString h & vbLf & ThePrivate & vbLf & ThePublic
End Sub

Public Sub SortAZ()
    If oModule.Component.CodeModule.countOfLines = 0 Then Exit Sub
    Dim coll        As New Collection: Set coll = Procedures
    Dim var         As Variant: ReDim var(1 To coll.count)
    Dim p           As aProcedure
    Dim StartLine   As Long
    Dim i           As Long
    For i = 1 To coll.count
        Set p = coll(i)
        var(i) = p.Name & "-" & p.KindAsLong
        If i = 1 Then StartLine = p.lines.first
    Next
    ArrayQuickSort var
    Dim S           As String: S = oModule.Header
    For i = LBound(var) To UBound(var)
        Set p = aProcedure.Initialize(oModule.WorkbookObject, oModule.Component, CStr(Split(var(i), "-")(0)), CLng(Split(var(i), "-")(1)))
        S = S & IIf(S <> "", vbNewLine, "") & p.code.All
    Next i
    oModule.CodeRemove
    oModule.Component.CodeModule.AddFromString S
End Sub

Sub Export(targetPath As String)
    If targetPath = "" Then Exit Sub    '<- FOR when called with FolderPicker
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim ap          As aProcedure
    Dim coll        As New Collection: Set coll = Procedures
    If coll.count > 0 Then
        For Each ap In coll
            'don't use this for classes !!! just export them as they are
            'because let/get/set use same name
            TxtOverwrite targetPath & ap.Name & ".txt", ap.code.All
        Next
    End If
End Sub

Sub Update()
    Dim Procedure   As aProcedure
    For Each Procedure In Procedures
        Procedure.Dependencies.Update
    Next
End Sub

