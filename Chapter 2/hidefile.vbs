'hidefile.vbs
Dim objFileSystem, objFile

'check if the parameter count is a single parameter. 
' If not then show the command syntax
If WScript.Arguments.Count<>1 Then 
WScript.Echo " Syntax: " & WScript.ScriptName &  _   
"  FileName " & vbCrLf & _
         " Filename: the path of the file you wish to hide"
       WScript.Quit
End If

 'create FileSystem object
 Set objFileSystem = CreateObject("Scripting.FileSystemObject")
 On Error Resume Next
 'get the file specified in command line 
 Set objFile = objFileSystem.GetFile(WScript.Arguments(0))
 'check if error occured - file not found
 If Err Then 
     WScript.Echo "Error:File:'" & WScript.Arguments(0) & "' not found" 
 Else
     objFile.Attributes = 2
 End If
