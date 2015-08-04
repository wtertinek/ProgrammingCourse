Dim max, counter
max = CInt(InputBox("Please enter the number you want to count to:"))

MsgBox "Let's print all numbers from 1 to " & max
counter = 1

Do While counter <= 3
  MsgBox "counter=" & counter
  counter = counter + 1
Loop

MsgBox "Done"