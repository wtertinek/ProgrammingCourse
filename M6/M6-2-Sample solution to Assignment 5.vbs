Dim fso, folderName, folder, files, extension, extLength, count
Set fso = CreateObject("Scripting.FileSystemObject")
folderName = GetFolderName(fso, "Enter the name of the folder to scan:")
extension = LCase(InputBox("Enter the extension to scan for:"))

Set folder = fso.GetFolder(folderName)
Set files = folder.Files

count = 0
extLength = len(extension)

For Each file In files
  If Right(LCase(file.Name), extLength) = extension Then
    count = count + 1
  End If
Next

If count = 1 Then
  MsgBox "There is " & count & " file of type " & extension & " in folder " & folderName
Else
  MsgBox "There are " & count & " files of type " & extension & " in folder " & folderName
End If

Function GetFolderName(fso, message)
  Dim folderName
  
  Do
    folderName = InputBox(message)
    
    If folderName = "" Then
      WScript.Quit
    End If
  Loop Until fso.FolderExists(folderName)
  
  GetFolderName = folderName
End Function