Dim max
max = CInt(InputBox("Please enter the number you want to count to:"))

MsgBox "Let's print all numbers from 1 to " & max

For counter = 1 To max
  MsgBox "counter=" & counter
Next

MsgBox "Let's print all ODD numbers from 1 to " & max

For counter = 1 To max Step 2
  MsgBox "counter=" & counter
Next
