Const wbemFlagReturnImmediately  =16 
Const wbemFlagForwardOnly = 32

Dim objService, objWMIFiles, nResult, objFile

'get a reference to a WMI service
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}")

'get all files with .tmp extenstion on the C: drive
Set objWMIFiles = _
objService.ExecQuery ("Select Name From CIM_DataFILE Where " & _
                             "Drive='C:' And Extension='tmp'"”,, _
                               wbemFlagReturnImmediately + wbemFlagForwardOnly 
)
'loop through all files and attempt to delete them
For Each objFile In objWMIFiles
    nResult = objFile.Delete
    If nResult <> 0 Then
      WScript.Echo "*Unable to delete " & objFile.Name
    Else
      WScript.Echo "Successfully deleted file:" & objFile.Name
    End If
Next
