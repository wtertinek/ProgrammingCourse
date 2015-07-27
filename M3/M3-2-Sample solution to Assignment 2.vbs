Dim firstNumber, secondNumber, operand, result

operand = InputBox("Choose the operation to perform (+, -, *, /, sqr, sin, cos, tan):")

If operand = "sqr" Or operand = "sin" Or operand = "cos" Or operand = "tan" Then
  firstNumber = CInt(InputBox("Please enter the number:"))
  
  If operand = "sqr" Then
    result = sqr(firstNumber)
  ElseIf operand = "sin" Then
    result = sin(firstNumber)
  ElseIf operand = "cos" Then
    result = cos(firstNumber)
  ElseIf operand = "tan" Then
    result = tan(firstNumber)
  End If
  
  MsgBox operand & "(" & firstNumber & ") = " & result
Else
  firstNumber = CInt(InputBox("Please enter the first number:"))
  secondNumber = CInt(InputBox("Please enter the second number:"))

  If operand = "+" Then
    result = firstNumber + secondNumber
  ElseIf operand = "-" Then
    result = firstNumber - secondNumber
  ElseIf operand = "*" Then
    result = firstNumber * secondNumber
  ElseIf operand = "/" Then
    result = firstNumber / secondNumber
  Else
    MsgBox "Invalid operation"
    WScript.Quit
  End If
  
  MsgBox firstNumber & " " & operand & " " & secondNumber & " = " & result   
End If