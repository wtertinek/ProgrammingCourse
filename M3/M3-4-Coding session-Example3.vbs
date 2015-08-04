Dim number

Do
  number = InputBox("Please enter a number:")
'Change While Not to Until
Loop While Not IsNumeric(number)

MsgBox "Now the text can be converted to a number"
number = CInt(number)

MsgBox number & "+" & number & "=" & number + number