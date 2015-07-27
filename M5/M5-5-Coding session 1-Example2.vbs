Dim fso, input, drive
Set fso = CreateObject("Scripting.FileSystemObject")

Do
  input = InputBox("Enter a drive letter:")
  
  If input = "" Then
    WScript.Quit
  End If
Loop Until fso.DriveExists(input)

Set drive = fso.GetDrive(input)

If drive.IsReady Then
  MsgBox CInt(drive.FreeSpace / 10^6) & " MB free space available on drive " & drive.VolumeName
Else
  MsgBox "Drive is not ready"
End If