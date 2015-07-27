Dim numbers(), max

Do
  max = InputBox("Number of multiples:")
Loop Until IsNumeric(max)

max = CInt(max) - 1
ReDim numbers(max)

For i = 0 to max
  numbers(i) = 5 * (i + 1)
Next

MsgBox "The first " & max + 1 & " multiples of 5 are:"

For Each number in numbers
  MsgBox number
Next