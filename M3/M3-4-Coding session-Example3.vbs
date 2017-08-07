' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

Dim number

Do
  number = InputBox("Please enter a number:")
'Change While Not to Until
Loop While Not IsNumeric(number)

MsgBox "Now the text can be converted to a number"
number = CInt(number)

MsgBox number & "+" & number & "=" & number + number