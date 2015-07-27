Dim fso, scanFolder, targetFolder, textFile, folder, files
Set fso = CreateObject("Scripting.FileSystemObject")
scanFolder = GetFolderName(fso, "Enter the name of the folder to scan:")
targetFolder = GetFolderName(fso, "Enter the target folder:")

If fso.FileExists(targetFolder & "\" & "Files.txt") Then
  MsgBox "Error, file already exists"
else
  Set textFile = fso.CreateTextFile(targetFolder & "\" & "Files.txt")
  Set folder = fso.GetFolder(scanFolder)
  Set files = folder.Files
  
  For Each file In files
    textFile.WriteLine(file.Name)
  Next
  
  textFile.Close
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