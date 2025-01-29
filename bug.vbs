Function f(a, b)
  If IsEmpty(a) Then
    a = 0
  End If
  If IsEmpty(b) Then
    b = 0
  End If
  c = a + b
  f = c
End Function

MsgBox f(1, "") ' Output: 1
MsgBox f("", 2) ' Output: 2
MsgBox f("", "") ' Output: 0
MsgBox f(1, 2) ' Output: 3