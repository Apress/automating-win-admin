'regit.vbs
'checks for existence of specified DLL's and registers
'them if required
Const ForAppending = 8
Dim objShell, strComputer, strSource, strDest
Dim strLogFile,  objNetwork, strUser

Set objNetwork = CreateObject("WScript.Network")

'loop until user id is retrieved, required for Win9x
Do While strUser =""
 strUser = objNetwork.UserName
Loop

'set the log file name - ensure uniqueness by combining
'user id and date and time
strLogFile = "\\thor\e$\" & strUser & " " &  Month(date) & "-" & _
              Day(date) & "-" & Year(date) & " " & _
              Hour(time) & "_" & Minute(time) & "_" _
              & Second(time) & ".txt"

Set objShell =CreateObject("WScript.Shell")

'set destination directory depending on OS
If objShell.ExpandEnvironmentStrings("%OS%") = "Windows_NT" Then
  strDest = objShell.ExpandEnvironmentStrings("%windir%") _
                & "\system32\"
Else
  strDest = objShell.ExpandEnvironmentStrings("%windir%") _
                & "\system\"
End If

'get the source directory to find the components to register
strSource = objShell.ExpandEnvironmentStrings("%LDIR%")

CheckRegister "regobj.dll", False


'CheckRegister
'Checks for existence of specified file and copies and registers it
'if it does not exist.
'Parameters
'strFile   Name of file to check for
'bReplace  Boolean value. If True then file will be updated if newer
'          version exists
Sub CheckRegister(strFile, bReplace)

Dim strPath
Dim objFSO
Dim bRegister, strDstVer, strSrcVer

strComputer = ""

strDest = strDest & strFile
strSource = strSource & strFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
 bRegister = False
 
 'check if specified file exists
 If Not objFSO.FileExists(strDest) Then
    objFSO.CopyFile strSource, strDest
    bRegister = True
 Else
  'file exists.. replace?
  If bReplace Then
  
   On Error Resume Next
   'attempt to get file version of specified files
   strSrcVer = objFSO.GetFileVersion(strSource)
   strDstVer = objFSO.GetFileVersion(strDest)
   'error ocurred.. unable to get version, use file size instead
   If Err Then
    strSrcVer = objFSO.GetFile(strSource).Size
    strDstVer = objFSO.GetFile(strDest).Size
   End If
   Err.Clear
   'check if the destination file version is less than source
    If Val(strSrcVer) > Val(strDstVer) Then
     'copy over existing file
     objFSO.CopyFile strSource, strDest
      'error copying file?
      If Err Then
       'error copying source to destination..
         LogIt strLogFile, "Error copying file " & strSource & _
               " to " & strDest & " on computer " & strComputer _
               & vbCrLf & Err.Description & " " & Err
        CheckRegister = False
       Exit Sub
      End If
     bRegister = True
    Else
     'more recent version of file, exit function
     LogIt strLogFile, "File " & strSource & _
          " is more recent than " & strDest & " on computer " & _
          strComputer
     Exit Sub
    End If
  End If
 
 End If
 
 'register the file?
 If bRegister Then
    objShell.Run strSource & "regsvr32 /s " & strDest, 0, True
 End If
 
End Sub

'Procedure Logfile
'Logs text to specified file
'Parameters
'strFile   Path to log file
'strMsg    Message to log
Sub LogIt(strFile, strMsg)
Dim objTS, objFSO

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTS = objFSO.OpenTextFile(strFile, ForAppending, True)
objTS.WriteLine Now & " " & strMsg

objTS.Close
End Sub
