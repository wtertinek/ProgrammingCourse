Dim countFrom, countTo, steps

countFrom = InputBox("Please enter the number to start from:")
countTo = InputBox("Please enter the number to count to:")
steps = InputBox("Please enter the step width:")

For number = countFrom To countTo Step steps
  MsgBox number
Next