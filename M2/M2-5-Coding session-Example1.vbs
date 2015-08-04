Dim answer, number
Number = CInt(InputBox("Please enter a number"))

'We assign the result of the comparison number = 5 to the variable answer
answer = number = 5

if answer then
  MsgBox "You entered 5"
else
  MsgBox "You entered a number different from 5"
end if