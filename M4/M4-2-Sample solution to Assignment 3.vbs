' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

Dim countFrom, countTo, steps

countFrom = InputBox("Please enter the number to start from:")
countTo = InputBox("Please enter the number to count to:")
steps = InputBox("Please enter the step width:")

For number = countFrom To countTo Step steps
  MsgBox number
Next