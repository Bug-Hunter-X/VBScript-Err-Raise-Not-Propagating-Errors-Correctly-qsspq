Function MyFunction(param1)
  If IsEmpty(param1) Then
    Err.Raise vbError, , "Parameter cannot be empty"
  End If
  ' ... rest of the function
End Function