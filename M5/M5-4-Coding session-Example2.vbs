Dim sum

'Add this to see control flow
'MsgBox "Before the function call"
sum = GetSum(5, 7)
MsgBox "The result is " & sum
'MsgBox "After the function, the result is " & sum

Function GetSum(sum1, sum2)
  'Add this to see control flow
  'MsgBox "In the function"
  GetSum = sum1 + sum2
End Function