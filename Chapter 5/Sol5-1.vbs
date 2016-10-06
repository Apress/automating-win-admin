Dim objFSO, objDrive
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objDrive = objFSO.GetDrive("A")
'check if drive is ready, 
If objDrive.IsReady Then
 'drive is ready, do something..
End If
