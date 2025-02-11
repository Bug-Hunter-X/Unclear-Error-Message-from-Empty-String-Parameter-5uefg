Function MyFunction(param1)
  ' Check for empty parameter
  If IsEmpty(param1) Or param1 = "" Then
    ' Raise a more descriptive error
    Err.Raise vbError, , "Error: The parameter 'param1' cannot be empty or null." 
    Exit Function 'Important to exit after raising a custom error
  End If
  'Rest of the function code
  ' ...
End Function