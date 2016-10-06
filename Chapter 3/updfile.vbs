'updfiles.vbs
'copy files upon logon
Dim strPath, objFSO, strVal, objShell, strCopyPath

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

strPath = objShell.ExpandEnvironmentStrings("%LDIR%")
On Error Resume Next
'get reference to registry flag
strVal = _ 
   objShell.RegRead("HKCU\SOFTWARE\WSHUpdates\DeskTopShortcutsUpdate1")
'if registry key didn't exist, then copy files to desktop
If IsEmpty(strVal) Then
 strCopyPath = objShell.SpecialFolders("Desktop") & "\"
 objFSO.CopyFile strPath & "E-mail Policy.doc.lnk", strCopyPath, True
 objFSO.CopyFile strPath & "Phone List.doc.lnk", strCopyPath, True
 'update registry entry to reflect the operation has been performed
 objShell.RegWrite "HKCU\SOFTWARE\WSHUpdates\DeskTopShortcutsUpdate1", _
                   Date 
End If
strVal = Empty
'get reference to font update flag under local machine
strVal = objShell.RegRead("HKLM\SOFTWARE\WSHUpdates\TreFontUpd")

'if registry key didn't exist, then copy files to desktop
If IsEmpty(strVal) Then
  strCopyPath = objShell.SpecialFolders("Fonts") & "\"
  objFSO.CopyFile strPath & "Trebucbd.ttf", strCopyPath, True
  objShell.RegWrite "HKLM\SOFTWARE\WSHUpdates\TreFontUpd", Date
End If
