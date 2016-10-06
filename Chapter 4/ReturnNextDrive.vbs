'ReturnNextDrive.vbs
Function ReturnNextDrive()
Dim nF, objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'loop through drives starting from D:
For nF = ASC("d")  To ASC("z")
'if drive doesn't exist, it's available
 If Not objFSO.DriveExists(Chr(nF) & ":") Then
    ReturnNextDrive = Chr(nF) & ":"
    Exit Function
 End If
Next
End Function
