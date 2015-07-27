Dim number1, number2, operand

number1 = GetNumber("Please enter the first number")
operand = GetOperand()
number2 = GetNumber("Please enter the second number")

If operand = "+" Then
  ShowResult number1, "+", number2, number1 + number2
ElseIf operand = "-" Then
  ShowResult number1, "-", number2, number1 - number2
ElseIf operand = "*" Then
  ShowResult number1, "*", number2, number1 * number2
ElseIf operand = "/" Then
  ShowResult number1, "/", number2, number1 / number2
End If

Function GetNumber(text)
  Dim input

  Do
    input = InputBox(text)
    
    If input = "" Then
      WScript.Quit
    End If
  Loop Until IsNumeric(input)
  
  GetNumber = CInt(input)
End Function

Function GetOperand()
  Dim input

  Do
    input = InputBox("Choose the operation to perform (+, -, *, /):")

    If input = "" Then
      WScript.Quit
    End If
  Loop Until input = "+" Or input = "-" Or input = "*" Or input = "/"
  
  GetOperand = input
End Function

Sub ShowResult(number1, operand, number2, result)
  MsgBox number1 & operand & number2 & " = " & result
End Sub