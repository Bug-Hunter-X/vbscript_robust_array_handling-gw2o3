Function f(a)
  If IsEmpty(a) Then Exit Function
  For i = 1 To UBound(a)
    If a(i) = 0 Then
      f = f + 1
    End If
  Next
End Function

'This function should count the number of zeros in an array.
'However, it will produce an error if the array is empty or contains non-numeric values.
'This is because UBound will return an error if the array is empty,
and attempting to access a non-numeric value in the array with a(i) will also generate an error.