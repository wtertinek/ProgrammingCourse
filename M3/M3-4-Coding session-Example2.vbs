' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

Dim max, counter
max = CInt(InputBox("Please enter the number you want to count to:"))

MsgBox "Let's print all numbers from 1 to " & max
counter = 1

Do While counter <= 3
  MsgBox "counter=" & counter
  counter = counter + 1
Loop

MsgBox "Done"