Attribute VB_Name = "z_zTest"

Option Explicit

Sub vbTargets()
    Dim oProject As VBProject
    Dim oComponent As VBComponent
    Dim oCodeModule As CodeModule
    Dim oCodePane As CodePane
    Dim oDesigner As Object
    Set oProject = ThisWorkbook.VBProject
    Dim i As Long
    Dim S As String
    Dim val
    Dim y As Long
    Dim ctrl As MSForms.control
        S = S & IIf(S <> "", vbNewLine, "") & oProject.fileName
        For Each oComponent In oProject.VBComponents
            S = S & IIf(S <> "", vbNewLine, "") & "Name" & vbTab & oComponent.Name
            S = S & IIf(S <> "", vbNewLine, "") & "Type" & vbTab & oComponent.Type
            If oComponent.Type = vbext_ct_MSForm Then
                Set oDesigner = oComponent.Designer
                S = S & IIf(S <> "", vbNewLine, "") & "Controls:"
                i = 0
                For Each ctrl In oDesigner.Controls
                    i = i + 1
                    S = S & IIf(S <> "", vbNewLine, "") & i & vbTab & ctrl.Name & vbTab & TypeName(ctrl)
                Next
                With oComponent.DesignerWindow
                    
                End With
                i = 0
                S = S & IIf(S <> "", vbNewLine, "") & "Properties"
                For i = 1 To oComponent.Properties.count
                On Error Resume Next
                    S = S & IIf(S <> "", vbNewLine, "") & i & vbTab & oComponent.Properties(i).Name & vbTab & oComponent.Properties(i).value
                On Error GoTo 0
                Next
            End If
            S = S & IIf(S <> "", vbNewLine, "") & String(30, "-")
        Next
    Debug.Print S
    'split(s,String(30, "-"))(30)
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
