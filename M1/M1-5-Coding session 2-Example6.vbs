Dim a, b, result

' First use this code
a = InputBox("Enter first number:")
b = InputBox("Enter second number:")

' Then this
'MsgBox TypeName(a)
'a = CInt(InputBox("Enter first number:"))
'b = CInt(InputBox("Enter second number:"))

result = a + b

'If a equals 5 and b equals 7, the the message is
' "5 + 7 = 12"
MsgBox a & " + " & b & " = " & result