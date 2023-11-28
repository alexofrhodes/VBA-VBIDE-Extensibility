VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'https://codereview.stackexchange.com/questions/135926/managing-a-programmatically-accessible-stack-trace


''Example
'''''''''
'Sub DoSomething(ByVal value1 As Integer, ByVal value2 As Integer, ByVal value3 As String)
'    Const C_PROCEDURE_NAME = "DoSomething"
'    CallStack.Push C_MODULE_NAME, C_PROCEDURE_NAME, value1, value2, value3
'    TestSomethingElse value1
'NormalExit:
'    CallStack.Pop
'End Sub
'
'Private Sub TestSomethingElse(ByVal value1 As Integer)
'    Const C_PROCEDURE_NAME = "TestSomethingElse"
'    CallStack.Push C_MODULE_NAME, C_PROCEDURE_NAME, value1
'    On Error GoTo ErrorHandler
'
'    Debug.Print value1 / 0
'
'NormalExit:
'    CallStack.Pop
'    Exit Sub
'ErrorHandler:
'    CallStack.PrintErrorInfo
'    Resume NormalExit
'End Sub

Option Explicit
Private frames As New Collection

Public Sub Push(ByVal Module As String, ByVal member As String, ParamArray parameterValues() As Variant)
Attribute Push.VB_Description = "Pushes a new stack frame onto the call stack. Call once at the entry point of each procedure to trace."
    Dim values() As Variant
    values = parameterValues
    frames.Add CallStackFrame.Create(Module, member, values)
End Sub

Public Function Pop() As CallStackFrameInterface
Attribute Pop.VB_Description = "Removes the last stack frame from the top of the stack. Call once at the exit point of each traced procedure."
    Set Pop = Peek
    frames.Remove frames.count
End Function

Public Function Peek() As CallStackFrameInterface
Attribute Peek.VB_Description = "Returns the top-most stack frame."
    Set Peek = frames(frames.count)
End Function

Public Property Get count() As Long
Attribute count.VB_Description = "Gets the depth of the call stack."
    count = frames.count
End Property

Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a String containing the stack trace."
    Dim result As String
    Dim index As Long
    For index = frames.count To 1 Step -1
        result = result & "at " & frames(index).ToString & IIf(index = 1, vbNullString, vbNewLine)
    Next
    ToString = result
End Function

Public Function ErrorInfo() As String
    ErrorInfo = "Runtime error " & Err.Number & ": " & Err.Description & vbNewLine & CallStack.ToString
End Function

Public Sub PrintErrorInfo()
    Debug.Print ErrorInfo
End Sub

Public Sub LogErrorInfo()
    LogFile.Log ErrorInfo
End Sub
