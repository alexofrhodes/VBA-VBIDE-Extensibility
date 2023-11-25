VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedureInject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize          As aProcedureInject  :
' Get ObjectsReleaseText  As String            :
' Sub BodyAfterComments                        :
' Sub BodyBottom                               :
' Sub BodyTop                                  :
' Sub Description                              :
' Sub HeaderBottom                             :
' Sub HeaderTop                                :
' Sub Modification                             :
' Sub ObjectsReleaseAtEnd                      :
' Sub ObjectsReleaseHere                       :
' Sub Template                                 :
' Sub TemplateObject                           :
' Sub Timer                                    :
' Sub test                                     :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit

Private oProcedure As aProcedure

Public Function Initialize(Procedure As aProcedure) As aProcedureInject
    Set oProcedure = Procedure
    Set Initialize = Me
End Function

Public Sub HeaderTop(this As String, SkipIfExists As Boolean)
    If SkipIfExists Then If InStr(1, oProcedure.Code.Body, this) > 0 Then Exit Sub
    oProcedure.Parent.CodeModule.InsertLines oProcedure.Code.lines.first, vbNewLine & this
End Sub

Public Sub HeaderBottom(this As String, SkipIfExists As Boolean)
    If SkipIfExists Then If InStr(1, oProcedure.Code.Body, this) > 0 Then Exit Sub
    oProcedure.Parent.CodeModule.InsertLines oProcedure.Code.lines.FirstOfDeclaration, this
End Sub

Public Sub BodyTop(this As String, SkipIfExists As Boolean)
    If SkipIfExists Then If InStr(1, oProcedure.Code.Body, this) > 0 Then Exit Sub
    oProcedure.Parent.CodeModule.InsertLines oProcedure.Code.lines.FirstOfBody, this
End Sub

Public Sub BodyAfterComments(this As String, SkipIfExists As Boolean)
    If SkipIfExists Then If InStr(1, oProcedure.Code.Body, this) > 0 Then Exit Sub
    oProcedure.Parent.CodeModule.InsertLines oProcedure.Code.lines.FirstOfBodyAfterComments, this
End Sub

Public Sub BodyBottom(this As String, SkipIfExists As Boolean)
    If SkipIfExists Then If InStr(1, oProcedure.Code.Body, this) > 0 Then Exit Sub
    oProcedure.Parent.CodeModule.InsertLines oProcedure.Code.lines.LastOfBody + 1, this
End Sub


Public Sub Template(Optional addErrorHandling As Boolean = True)
    Dim newCode As String
    newCode = Join( _
                Array( _
                    oProcedure.Code.Declaration, _
                    NewProcedureTemplate(False, addErrorHandling), _
                    oProcedure.oModule.CodeModule.lines(oProcedure.lines.last, 1)), _
                vbLf)
    oProcedure.Format.Replace newCode
    Modification
End Sub

Public Sub TemplateObject(Optional addErrorHandling As Boolean = True)
    Dim newCode As String
    newCode = Join( _
                Array( _
                    oProcedure.Code.Declaration, _
                    NewProcedureTemplate(True, addErrorHandling), _
                    oProcedure.oModule.CodeModule.lines(oProcedure.lines.last, 1)), _
                vbLf)
    oProcedure.Format.Replace newCode
    Modification
End Sub

Private Function NewProcedureTemplate(ReturnsObject As Boolean, addErrorHandling As Boolean) As String
    Const C_Tab = vbTab
    Dim isFunction As Boolean
    Select Case oProcedure.KindAsString
    Case "Function", "Get": isFunction = True
    End Select
    
    Dim Q As String: Q = """"
    
    Dim am As New aModule: Set am = am.Initialize(oProcedure.Parent)
    If Not am.HeaderContains("C_MODULE_NAME") Then
        oProcedure.Parent.CodeModule.AddFromString "Private Const C_MODULE_NAME = " & Q & am.Name & Q
    End If
    
    Dim S As String, msg As String
    S = Join( _
            Array( _
                "'=== Constants ===", _
                C_Tab & "Const C_PROCEDURE_NAME = " & Q & oProcedure.Name & Q, _
                "", _
                "'=== Variables ==="), _
            vbLf)
    If isFunction Then S = Join(Array(S, C_Tab & "Dim Output As " & oProcedure.returnType, ""), vbLf)
    
    S = Join( _
            Array( _
                S, _
                "'=== Initialize ==="), _
            vbLf)
    
    msg = "Now & vbTab & " & Q & "Start:" & Q & "& vbTab & C_PROCEDURE_NAME"
    S = Join( _
            Array( _
                S, _
                C_Tab & "debug.print " & msg, _
                C_Tab & "LogFile.log " & msg), _
            vbLf)
    
    Dim args As String: args = oProcedure.arguments.AsSeen
    
    S = Join( _
            Array( _
                S, _
                C_Tab & "CallStack.Push C_MODULE_NAME, C_PROCEDURE_NAME" & IIf(args <> "", ", " & Q & args & Q, ""), _
                IIf(addErrorHandling, C_Tab & "On error goto ErrorHandler", ""), _
                "", _
                "'=== Code ===", _
                oProcedure.Code.Body, _
                "", _
                "", _
                "'=== Terminate ===", _
                "NormalExit:"), _
            vbLf)

    If isFunction Then S = Join(Array(S, C_Tab & "" & IIf(ReturnsObject, "Set ", "") & oProcedure.Name & " = output"), vbLf)
    
    msg = VBA.Replace(msg, "Start:", "End:", , 1, vbTextCompare)
    S = Join( _
            Array( _
                S, _
                C_Tab & "debug.print " & msg, _
                C_Tab & "LogFile.log " & msg, _
                C_Tab & "CallStack.Pop", _
                "", _
                "'=== Clean Up ===", _
                "'aprocedure.active.inject.ObjectsReleaseHere", _
                ObjectsReleaseText, _
                "", _
                C_Tab & "Exit " & oProcedure.KindAsString, _
                ""), _
            vbLf)
    If addErrorHandling Then
        S = Join( _
                Array( _
                    S, _
                    "ErrorHandler:", _
                    C_Tab & "CallStack.PrintErrorInfo", _
                    C_Tab & "CallStack.LogErrorInfo", _
                    C_Tab & "ERR.CLEAR", _
                    C_Tab & "ERR.Raise 999,C_PROCEDURE_NAME,CALLSTACK.ErrorInfo"), _
                vbLf)
    End If
    NewProcedureTemplate = S
    

'note'
''''''
'    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
'            "Error Number: " & Err.Number & vbCrLf & _
'            "Error Source: TxtOverwrite" & vbCrLf & _
'            "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
End Function

Sub Description()
    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab2 & vbTab2
    Const sFersLine As String = "'* Modified   :" & vbTab & "Date and Time" & vbTab2 & "Author" & vbTab4 & "Description" & vbCrLf

    Dim txtName     As String: txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")

    Dim txtContacts As String: txtContacts = AUTHOR_EMAIL
    If txtContacts <> vbNullString Then txtContacts = "'* Contacts   :" & vbTab & txtContacts & vbCrLf

    Dim txtCopyright As String: txtCopyright = AUTHOR_COPYRIGHT
    If txtCopyright <> vbNullString Then txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf

    Dim txtOther    As String: txtOther = AUTHOR_OTHERTEXT
    If txtOther <> vbNullString Then txtOther = "'* Note       :" & vbTab & txtOther & vbCrLf

    Dim txtMedia    As String: txtMedia = "'* " & vbLf & AUTHOR_MEDIA
    Dim CurentCodePane As CodePane: Set CurentCodePane = oProcedure.Parent.CodeModule.CodePane
    Dim i           As Byte
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim sProc       As String: sProc = oProcedure.Name
    Dim nLine       As Long
    nLine = oProcedure.Code.lines.FirstOfDeclaration
    Dim sTemp       As String
    Dim sTime       As String
    Dim sType       As String
    Dim sProcDeclartion As String
    Dim sProcArguments As String

    Dim sUser       As String

    '* @TODO Created: 01-02-2023 08:20 Author: Anastasiou Alex
    '* @TODO if name is too long have to recalculate tab length

    sUser = "Alex"
    If sUser = vbNullString Then sUser = Environ("UserName")

    On Error Resume Next

    With CurentCodePane
        sTemp = VBA.Replace(String(90, "*"), "**", "* ")
        sTime = Format(Now, "dd-mm-yyyy hh:nn")

        For i = 0 To 4
            procKind = i
            sProcDeclartion = oProcedure.Code.DeclarationClean
            If sProcDeclartion <> vbNullString Then Exit For
        Next
        sProcArguments = AddStringParameterFromProcedureHeader(sProcDeclartion)
        sType = TypeProcedureComment(sProcDeclartion)
        sTemp = _
                "'" & sTemp & vbCrLf & _
                "'" & sType & vbTab & sProc & vbCrLf & _
                "'* Author     :" & vbTab & txtName & vbCrLf & _
                txtContacts & _
                txtCopyright & _
                txtOther & _
                "'* Purpose    :" & vbTab & vbCrLf & _
                txtMedia & _
                sFersLine & _
                "'* Created    :" & vbTab & sTime & vbTab & sUser & vbCrLf & _
                sProcArguments & _
                "'" & sTemp
        HeaderTop sTemp, False
    End With
End Sub

Private Function AddStringParameterFromProcedureHeader(ByVal sPocDeclartion As String) As String
    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab2 & vbTab2
    Dim sDeclaration As String
    sDeclaration = Right$(sPocDeclartion, Len(sPocDeclartion) - InStr(1, sPocDeclartion, "("))
    sDeclaration = Left$(sDeclaration, InStr(1, sDeclaration, ")") - 1)
    If sDeclaration = vbNullString Then Exit Function
    Dim arStr()     As String
    arStr = Split(sDeclaration, ",")
    Dim iMaxLen     As Byte
    iMaxLen = 0
    Dim iTempLen    As Byte
    Dim sTemp       As String
    Dim i           As Byte
    For i = 0 To UBound(arStr)
        iTempLen = Len(Trim$(arStr(i)))
        If iMaxLen < iTempLen Then iMaxLen = iTempLen
    Next i
    Dim numOfSpaces As Long
    numOfSpaces = iMaxLen - Len(Trim$("'* Argument(s):"))
    If numOfSpaces < 0 Then numOfSpaces = 0
    sDeclaration = "'*" & vbLf & "'* Argument(s):" & Space(numOfSpaces) & vbTab2 & "Description" & vbCrLf & "'*" & vbCrLf
    For i = 0 To UBound(arStr)
        sTemp = "'* " & Trim$(arStr(i)) & Space(iMaxLen - Len(Trim$(arStr(i)))) & " :"
        sDeclaration = sDeclaration & sTemp & vbCrLf
    Next i
    AddStringParameterFromProcedureHeader = sDeclaration & "'* " & vbCrLf
End Function

Private Function TypeProcedureComment(ByVal StrDeclarationProcedure As String) As String
    '@BlogPosted
    If StrDeclarationProcedure Like "*Sub*" Then
        TypeProcedureComment = "* Sub        :"
    ElseIf StrDeclarationProcedure Like "*Function*" Then
        TypeProcedureComment = "* Function   :"
    ElseIf StrDeclarationProcedure Like "*Property Set*" Then
        TypeProcedureComment = "* Prop Set   :"
    ElseIf StrDeclarationProcedure Like "*Property Get*" Then
        TypeProcedureComment = "* Prop Get   :"
    ElseIf StrDeclarationProcedure Like "*Property Let*" Then
        TypeProcedureComment = "* Prop Let   :"
    Else
        TypeProcedureComment = "* Un Type    :"
    End If
End Function

Sub Timer()
    If oProcedure.Code.Contains("StartTimer", True, True, False) Then Exit Sub
    BodyAfterComments "StartTimer " & """" & oProcedure.Name & """", True
    Sleep 200
    BodyBottom "EndTimer", True
End Sub

Sub test()
Dim x As Collection
Set x = New Collection
End Sub

Public Property Get ObjectsReleaseText() As String

    Dim i As Long
    Dim S As String, output As String, Keyword As String
    Dim vars As Collection
    Dim element
    Set vars = aProcedure.Variables.Items
    Dim av As aProcedureVariablesItem
    For i = oProcedure.lines.first To oProcedure.lines.last
        S = Trim(oProcedure.Parent.CodeModule.lines(i, 1))
        If (S Like "Set * = *" Or S Like "Dim*As New*") And Not S Like "*= Nothing*" Then
            Keyword = Split(S, " ")(1)
           
            For Each element In vars
                Set av = element
                If av.Name = Keyword Then
                    output = IIf(output <> "", output & vbNewLine, "") & "Set " & Keyword & " = Nothing"
                    Exit For
                End If
            Next
        End If
    Next
    If output = "" Then Exit Property
    Dim var: var = Split(output, vbNewLine)
    Dim sCode As String: sCode = oProcedure.Code.All
    Dim out, el
    ReDim out(0 To 0)
    For Each el In var
        If Not oProcedure.Code.Contains(CStr(el), True, True, False) Then
            If out(0) <> "" Then ReDim Preserve out(0 To UBound(out) + 1)
            out(UBound(out)) = el
        End If
    Next
    output = Space(4) & Join(out, vbNewLine & Space(4))
    ObjectsReleaseText = output
End Property
Public Sub ObjectsReleaseAtEnd()
    Dim this As String: this = ObjectsReleaseText
    If this <> "" Then oProcedure.Parent.CodeModule.InsertLines oProcedure.lines.last, this
End Sub
Public Sub ObjectsReleaseHere()
    aCodeModule.Active.Insert Space(4) & Join(Split(ObjectsReleaseText, vbNewLine), vbNewLine & Space(4))
End Sub

Public Sub Modification(Optional Desc As String = "")
    oProcedure.CustomProperties.LastModified = Format(Now, "yymmddhhnn")

    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab2 & vbTab2
    Dim sTime       As String: sTime = Format(Now, "dd-mm-yyyy hh:nn")
    Dim sProc       As String: sProc = oProcedure.Name

    Dim sUser       As String: sUser = "Alex"
    If sUser = vbNullString Then sUser = Environ("UserName")
    Const sUPDATE   As String = "'* Updated    :"
    Dim sFirstLine  As String
    sFirstLine = "'* Modified   :" & vbTab & "Date and Time" & vbTab2 & _
            "Author" & vbTab4 & "Description" & vbCrLf
    Dim sSecondLine As String
    sSecondLine = sUPDATE & vbTab & sTime & vbTab & sUser & vbTab4 & "(" & oProcedure.oModule.Name & aModule.Initialize(oProcedure.oModule).Extension & " > " & sProc & " : " & Desc & ")"

    If oProcedure.Code.ContainsInHeader("'* Created    :", True, False, False) Then
        Dim lineNo  As Long
        Dim i       As Long
        For i = oProcedure.lines.FirstOfHeader To oProcedure.lines.LastOfHeader
            If oProcedure.oModule.CodeModule.lines(i, 1) Like "'* Created    :*" Then
                lineNo = i
                Exit For
            End If
        Next
        If oProcedure.oModule.CodeModule.lines(lineNo + 1, 1) Like sUPDATE & "*" Then
            Do
                lineNo = lineNo + 1
            Loop While oProcedure.oModule.CodeModule.lines(lineNo + 1, 1) Like sUPDATE & "*"
        End If
        oProcedure.oModule.CodeModule.InsertLines lineNo + 1, sSecondLine
        Exit Sub
    End If

    Dim nLine       As Long: nLine = oProcedure.lines.FirstOfDeclaration
    If nLine < 4 Then
        sSecondLine = vbLf & sFirstLine & sSecondLine
    ElseIf Not oProcedure.oModule.CodeModule.lines(nLine - 2, 1) Like sUPDATE & "*" Then
        sSecondLine = vbLf & sFirstLine & sSecondLine
    End If
    If Len(oProcedure.oModule.CodeModule.lines(nLine - 1, 1)) = 0 And nLine > 1 Then
        oProcedure.oModule.CodeModule.InsertLines nLine - 1, sSecondLine
    Else
        oProcedure.oModule.CodeModule.InsertLines nLine, sSecondLine & vbNewLine
    End If
End Sub

