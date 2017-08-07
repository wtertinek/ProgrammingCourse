' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

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