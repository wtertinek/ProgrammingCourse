Dim numbers(), count, max, tmp, sum

Do
  count = InputBox("Please enter the number of values:")
Loop Until IsNumeric(count)

count = Cint(count)
max = count - 1

ReDim numbers(max)

For idx = 0 To max
  Do
    tmp = InputBox("Please enter number " & idx + 1 & " of " & count)
  Loop Until IsNumeric(tmp)
  
  numbers(idx) = CInt(tmp)
Next

For Each number In numbers
  sum = sum + number
Next

MsgBox "The sum equals " & sum