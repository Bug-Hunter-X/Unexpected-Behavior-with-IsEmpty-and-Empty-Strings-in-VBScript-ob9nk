Function f(a, b)
  If a = "" Then
    a = 0
  ElseIf Not IsNumeric(a) Then
    Err.Raise 13, , "Invalid input for a: " & a
  End If
  If b = "" Then
    b = 0
  ElseIf Not IsNumeric(b) Then
    Err.Raise 13, , "Invalid input for b: " & b
  End If 
  c = CDbl(a) + CDbl(b)
  f = c
End Function

MsgBox f(1, "") ' Output: 1
MsgBox f("", 2) ' Output: 2
MsgBox f("", "") ' Output: 0
MsgBox f(1, 2) ' Output: 3
MsgBox f("abc", 2) 'Raises Error 13 
MsgBox f(1, "def") 'Raises Error 13