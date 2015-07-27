Dim fso, fileName, file
Set fso = CreateObject("Scripting.FileSystemObject")

Do
  fileName = InputBox("Enter the full path of the file to read:")
Loop Until fso.FileExists(fileName)

Set file = fso.OpenTextFile(fileName)

Do Until file.AtEndOfStream
  MsgBox file.ReadLine
Loop

file.Close