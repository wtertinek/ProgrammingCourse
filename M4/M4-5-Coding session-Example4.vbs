' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

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