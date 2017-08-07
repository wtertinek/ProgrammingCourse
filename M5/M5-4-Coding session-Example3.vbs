' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

Dim a, b, result

a = GetNumber("Enter first number:")
b = GetNumber("Enter second number:")
result = a + b
PrintResult result

Sub PrintResult(res)
  MsgBox "The result is: " & res
End Sub

Function GetNumber(message)
  Dim tmp
  Do
    tmp = InputBox(message)
  Loop Until IsNumeric(tmp)
  
  GetNumber = CInt(tmp)
End Function