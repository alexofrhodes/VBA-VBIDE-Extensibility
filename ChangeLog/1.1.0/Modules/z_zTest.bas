Attribute VB_Name = "z_zTest"
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun sample2  As Collection     :
' Fun sample3  As String         :
' Sub GotoSelectedShapeOnAction  :
' Sub GotoShapeOnaction          :
' Sub PrintReturnTypeUnassigned  :
' Sub ShowInNotepad              :
' Sub sample1                    :
' Sub vbTargets                  :
'--------------------------------------------------
'@EndFolder ListOfProcedures

Option Explicit


Sub PrintReturnTypeUnassigned()
    Dim p As aProcedure
    For Each p In aModule.Active.Procedures.PublicProcedures
        If (p.returnType = "Variant" And Not p.code.Contains("set " & p.Name, True, False, False)) Then dp p.Name
    Next
End Sub

Sub vbTargets()
    Dim oProject As VBProject
    Dim oComponent As VBComponent
    Dim oCodeModule As CodeModule
    Dim oCodePane As CodePane
    Dim oDesigner As Object
    Set oProject = ThisWorkbook.VBProject
    Dim i As Long
    Dim s As String
    Dim val
    Dim y As Long
    Dim ctrl As MSForms.control
    s = s & IIf(s <> "", vbNewLine, "") & oProject.fileName
    For Each oComponent In oProject.VBComponents
        s = s & IIf(s <> "", vbNewLine, "") & "Name" & vbTab & oComponent.Name
        s = s & IIf(s <> "", vbNewLine, "") & "Type" & vbTab & oComponent.Type
        If oComponent.Type = vbext_ct_MSForm Then
            Set oDesigner = oComponent.Designer
            s = s & IIf(s <> "", vbNewLine, "") & "Controls:"
            i = 0
            For Each ctrl In oDesigner.Controls
                i = i + 1
                s = s & IIf(s <> "", vbNewLine, "") & i & vbTab & ctrl.Name & vbTab & TypeName(ctrl)
            Next
            With oComponent.DesignerWindow

            End With
            i = 0
            s = s & IIf(s <> "", vbNewLine, "") & "Properties"
            For i = 1 To oComponent.Properties.count
                On Error Resume Next
                s = s & IIf(s <> "", vbNewLine, "") & i & vbTab & oComponent.Properties(i).Name & vbTab & oComponent.Properties(i).value
                On Error GoTo 0
            Next
        End If
        s = s & IIf(s <> "", vbNewLine, "") & String(30, "-")
    Next
    Debug.Print s

End Sub

Sub GotoSelectedShapeOnAction()
Dim shapeCount  As Long
On Error Resume Next
shapeCount = Selection.ShapeRange.count
On Error GoTo 0
If shapeCount <> 1 Then Exit Sub
Dim shp         As Shape
Set shp = ActiveSheet.Shapes(Selection.Name)
GotoShapeOnaction ActiveSheet.Shapes(Selection.Name)
End Sub


Sub GotoShapeOnaction(shp As Shape)
    Dim Procedure   As String
    Procedure = shp.OnAction
    If Procedure = "" Then Exit Sub
    On Error GoTo ErrorHandler
    aProcedure.Initialize(ActiveWorkbook, , Procedure).Activate
    Exit Sub
ErrorHandler:
    MsgBox Procedure & " not found"
End Sub


Sub ShowInNotepad(txt As String)
    Dim targetFile  As String
    targetFile = ThisWorkbook.path & "\tmp.txt"
    TxtOverwrite targetFile, txt
    FollowLink targetFile
End Sub
