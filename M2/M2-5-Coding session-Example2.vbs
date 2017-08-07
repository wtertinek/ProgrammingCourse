' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

Dim firstNumber, secondNumber

firstNumber = CInt(InputBox("Enter the first number:"))
secondNumber = CInt(InputBox("Enter the second number:"))

' Same as If firstNumber <> 5 Then
If Not firstNumber = 5 Then
  MsgBox "First number is not 5"
End If

If firstNumber = 5 And secondNumber = 5 Then
  'If branch
  MsgBox "Both numbers equal 5"
ElseIf firstNumber = 5 Or secondNumber = 5 Then
  'ElseIf branch
  MsgBox "At least one number equals 5"
Else
  'Else branch
  MsgBox "None of the numbers equals 5"
End If