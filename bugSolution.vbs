Function MyFunction(param1, param2)
  If VarType(param1) = vbEmpty Then
    param1 = ""
  End If
  If VarType(param2) = vbEmpty Then
    param2 = 0
  ElseIf VarType(param2) <> vbInteger And VarType(param2) <> vbLong Then
    param2 = 0 ' Handle non-numeric input as 0
  End If
  ' ... rest of the function ...
End Function