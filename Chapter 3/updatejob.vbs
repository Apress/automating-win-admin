'updatejob.vbs
'copy job file from logon location to Tasks folder 
Const WindowsFolder = 0
Dim objShell, strPath, objFSO, strVal

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

Set objShell = CreateObject("WScript.Shell")
'set the directory to find the job file to copy
strPath = "\\odin\jobfiles\"

On Error Resume Next
'get reference to registry flag
strVal = _ 
    objShell.RegRead("HKCU\SOFTWARE\WSHUpdates\AddJob1")
'if registry key didn't exist, then copy files to Windows Tasks folder
If IsEmpty(strVal) Then
 strCopyPath = objFSO.GetSpecialFolder(WindowsFolder) & "\Tasks\"

 objFSO.CopyFile strPath & "maintenance.job", strCopyPath , True
'update registry entry to reflect the operation has been performed
 objShell.RegWrite "HKCU\SOFTWARE\WSHUpdates\AddJob1", Date 
End If
