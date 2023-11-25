VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProjectReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@FOLDER ListOfProcedures
'--------------------------------------------------
' Fun Initialize  As aProjectReferences  :
' Sub AddFromFile                        :
' Sub AddFromGUID                        :
' Sub AddScriptControl                   :
' Sub AddVBIDE                           :
' Sub Export                             :
' Sub ImportReferences                   :
' Sub RemoveByDescription                :
' Sub RemoveByGUID                       :
' Sub RemoveByName                       :
' Sub ToSheet                            :
'--------------------------------------------------
'@EndFolder ListOfProcedures
Option Explicit

Private oProject As aProject

Public Function Initialize(Project As aProject) As aProjectReferences
    Set oProject = Project
    Set Initialize = Me
End Function


Sub ToSheet()
    Dim ws          As Worksheet
    Set ws = CreateOrSetSheet("References", oProject.WorkbookObject)
    ws.Cells.clear
    ws.Cells(1, 1).value = "Reference Description"
    ws.Cells(1, 2).value = "GUID"
    ws.Cells(1, 3).value = "Path"
    ws.Cells(1, 4).value = "Version"
    Dim myRef       As Reference
    Dim refs        As VBIDE.REFERENCES
    Set refs = oProject.Project.REFERENCES
    Dim i           As Long
    i = 2
    For Each myRef In refs
        ws.Cells(i, 1) = IIf(myRef.Description <> "", myRef.Description, myRef.Name)
        ws.Cells(i, 2) = myRef.GUID
        ws.Cells(i, 3) = myRef.fullPath
        ws.Cells(i, 4) = myRef.major & "." & myRef.minor
        i = i + 1
    Next myRef
End Sub

Sub AddFromFile(FilePath As String)
    oProject.Project.REFERENCES.AddFromFile FilePath
End Sub

Sub AddFromGUID(GUID As String, major As String, minor As String)
    oProject.Project.REFERENCES.AddFromGUID GUID, major, minor
End Sub

Sub RemoveByName(RefName As String)
    Dim oRef        As Reference
    Dim oRefs       As REFERENCES: Set oRefs = oProject.Project.REFERENCES
    For Each oRef In oRefs
        If oRef.Name = RefName Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Sub RemoveByDescription(RefDescription As String)
    Dim oRef        As Reference
    Dim oRefs       As REFERENCES: Set oRefs = oProject.Project.REFERENCES
    For Each oRef In oRefs
        If oRef.Description = RefDescription Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Sub RemoveByGUID(refGUID As String)
    Dim oRefs       As REFERENCES: Set oRefs = oProject.Project.REFERENCES
    Dim oRef        As Reference
    For Each oRef In oRefs
        If oRef.GUID = refGUID Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Public Sub Export(FolderPath As String)
    Dim output      As String
    Dim oRef        As Reference
    For Each oRef In oProject.Project.REFERENCES
        output = output & IIf(output <> "", vbNewLine, "") & Join( _
                                                                Array( _
                                                                    oRef.Name, _
                                                                    oRef.GUID, _
                                                                    oRef.major, _
                                                                    oRef.minor), _
                                                                vbTab)
    Next
    TxtOverwrite VBA.Replace(FolderPath & "\", "\\", "\") & _
                 Split(oProject.WorkbookObject.Name, ".")(0) & _
                 "_References.txt", _
                 output
End Sub


Public Sub ImportReferences(FilePath As String)
    Dim FSO         As New Scripting.FileSystemObject
    With FSO.OpenTextFile(FilePath, ForReading, True)
        Dim Line    As Long
        Do While Not .AtEndOfStream
            Dim values As Variant
            values = Split(.ReadLine, vbTab)
            On Error Resume Next
            oProject.Project.REFERENCES.AddFromGUID values(1), values(2), values(3)
        Loop
    End With
End Sub

Public Sub AddVBIDE()
    On Error Resume Next
    oProject.Project.REFERENCES.AddFromGUID "{0002E157-0000-0000-C000-000000000046}", 5, 3
End Sub

Sub AddScriptControl()
    On Error Resume Next
    oProject.Project.REFERENCES.AddFromGUID "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}", 1, 0
End Sub

