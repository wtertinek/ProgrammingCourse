' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

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
