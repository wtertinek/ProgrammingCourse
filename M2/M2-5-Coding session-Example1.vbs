' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

Dim answer, number
Number = CInt(InputBox("Please enter a number"))

'We assign the result of the comparison number = 5 to the variable answer
answer = number = 5

if answer then
  MsgBox "You entered 5"
else
  MsgBox "You entered a number different from 5"
end if