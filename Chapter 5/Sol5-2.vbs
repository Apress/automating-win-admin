Dim objFSO, objDrive
Set objFSO = CreateObject("Scripting.FileSystemObject")
For Each objDrive In objFSO.Drives
   'check if drive is ready, 
   If objDrive.IsReady Then
      Wscript.Echo objDrive.DriveLetter & " is " & _
      Fix(((objDrive.TotalSize - objDrive.FreeSpace) _
            / objDrive.TotalSize) * 100) & "% used"
   Else
      Wscript.Echo objDrive.DriveLetter & " is not ready"
   End If
Next
