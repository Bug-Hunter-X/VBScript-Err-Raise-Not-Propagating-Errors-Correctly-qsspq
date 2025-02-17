Function MyFunction(param1)
  On Error Resume Next ' Enable error handling within this function
  If IsEmpty(param1) Then
    Err.Raise vbError, , "Parameter cannot be empty"
  End If
  On Error GoTo 0 ' Disable error handling
  ' ... rest of the function
End Function

Sub CallMyFunction()
  On Error GoTo ErrorHandler
  Dim result
  result = MyFunction("")
  Exit Sub
ErrorHandler:
  MsgBox "Error: " & Err.Description
End Sub