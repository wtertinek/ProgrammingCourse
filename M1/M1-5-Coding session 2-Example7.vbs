Dim a, result

'We let the user enter a number and raise e to the power of it
a = CDbl(InputBox("Enter a number to calculate e^x:"))
MsgBox "e^" & a & " = " & Round(Exp(a), 2)

'We let the user enter a number and take the square root of it
a = CDbl(InputBox("Enter a positive integer to calculate the square root:"))
MsgBox "The square root of " & a & " is " & Round(Sqr(a), 2)

'We let the user enter a number and take the absolute value of it
a = CDbl(InputBox("Enter a number:"))
MsgBox "The absolute value of " & a & " is " & Abs(a)

'We let the user enter an angle (in radians) and calculate the sine from it
a = CDbl(InputBox("Enter an angle in radians to calculate the sine:"))
MsgBox "Sine of " & a & " is " & Round(Sin(a), 2)
