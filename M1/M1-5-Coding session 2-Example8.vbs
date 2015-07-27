Dim text1, text2, text3, subString, position

text1 = InputBox("Enter a text")
text2 = InputBox("Enter another text")

text3 = text1 & " " & text2

MsgBox "text1 + text2 is " & Len(text3) & " characters long: " & text3

'Lets check if the value of subString occurs in text3
subString = InputBox("Enter a sub string to look for:")
position = InStr(text3, subString)
MsgBox "The sub string '" & subString & "' starts at positionition " & position

'Lets display the n first characters
position = CInt(InputBox("Enter how many characters should be returned from the left:"))
MsgBox Left(text3, position)

'Lets display the n last characters
position = CInt(InputBox("Enter how many characters should be returned from the right:"))
MsgBox Right(text3, position)