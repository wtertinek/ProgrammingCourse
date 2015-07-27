Dim fileName, fso, file, text
Set fso = CreateObject("Scripting.FileSystemObject")

Do
  fileName = InputBox("Enter the full path of the file:")
  
  If fileName = "" Then
    WScript.Quit
  End If
Loop Until fso.FileExists(fileName)

Set file = fso.GetFile(fileName)

text = file.Path & vbCrLf & vbCrLf
text = text & "Created: " & file.DateCreated & vbCrLf
text = text & "Last accessed: " & file.DateLastAccessed & vbCrLf
text = text & "Last modified: " & file.DateLastModified

MsgBox text