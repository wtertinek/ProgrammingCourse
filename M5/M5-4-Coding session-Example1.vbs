' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

PrintSum 6, 8
PrintSum 2, 9

Sub PrintSum(number1, number2)
  Dim result
  result = number1 + number2
  MsgBox number1 & " + " & number2 & " = " & result
End Sub