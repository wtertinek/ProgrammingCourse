Dim values(), sum

ReDim values(0)
values(0) = 5

ReDim preserve values(1)
values(1) = 10

ReDim preserve values(2)
values(2) = 15

MsgBox "Value at index 0: " & values(0)
MsgBox "Value at index 1: " & values(1)
MsgBox "Value at index 2: " & values(2)

sum = values(0) + values(1) + values(2)
MsgBox "The sum of all value: " & sum