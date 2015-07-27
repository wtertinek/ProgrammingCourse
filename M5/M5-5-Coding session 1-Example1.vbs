Dim fso, sourceFile, targetFolder, targetFile
Set fso = CreateObject("Scripting.FileSystemObject")

Do
  sourceFile = InputBox("Enter the full path of the file to copy:")

  If sourceFile = "" Then
    WScript.Quit
  End If
Loop Until fso.FileExists(sourceFile)

Do
  targetFolder = InputBox("Enter the full path of the target folder:")

  If targetFolder = "" Then
    WScript.Quit
  End If
Loop Until fso.FolderExists(targetFolder)

targetFile = targetFolder & "\" & fso.GetFileName(sourceFile) & ".bak"

fso.CopyFile sourceFile, targetFile