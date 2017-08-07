' Wolfgang Tertinek, 2017
' This work is licensed under a Creative Commons Attribution 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

Dim values()

ReDim values(0)
values(0) = "Hello"
MsgBox "Number of array elements: " & UBound(values) + 1

ReDim Preserve values(1)
values(1) = "World"
MsgBox "Number of array elements: " & UBound(values) + 1

MsgBox values(0) & " " & values(1)

ReDim Preserve values(0)
MsgBox "Number of array elements: " & UBound(values) + 1

If IsArray(values) Then
  MsgBox "The variable named 'values' is an array."
End If