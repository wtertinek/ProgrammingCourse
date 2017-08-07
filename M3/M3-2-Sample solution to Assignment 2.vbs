' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

Dim firstNumber, secondNumber, operand, result

operand = InputBox("Choose the operation to perform (+, -, *, /):")

If operand = "+" Or operand = "-" Or operand = "*" Or operand = "/" Then
  firstNumber = GetNumber("Please enter the first number:")
  secondNumber = GetNumber("Please enter the second number:")

  If operand = "+" Then
    result = firstNumber + secondNumber
  ElseIf operand = "-" Then
    result = firstNumber - secondNumber
  ElseIf operand = "*" Then
    result = firstNumber * secondNumber
  ElseIf operand = "/" Then
    result = firstNumber / secondNumber
  End If
  
  MsgBox firstNumber & " " & operand & " " & secondNumber & " = " & result
Else
  MsgBox "Invalid operation"
End If

Function GetNumber(text)
  Dim input

  Do
    input = InputBox(text)
    
    If input = "" Then
      WScript.Quit
    End If
  Loop Until IsNumeric(input)
  
  GetNumber = CDbl(input)
End Function