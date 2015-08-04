Dim multiplier, numbers(5)

Do
  multiplier = InputBox("Please enter the multiplier:")
Loop Until IsNumeric(multiplier)

multiplier = CInt(multiplier)

For idx = 0 To 4
  numbers(idx) = (idx + 1) * multiplier
Next

MsgBox "The first value is:" & numbers(0)
MsgBox "The last value is:" & numbers(4)