'UpdateEnvironment
'Description
'updates autoexec.bat file with environment variables and sets 
'the environment variable
'Parameters:
'strVariablename of environment variable
'varValue       value to set
Sub UpdateEnvironment(strVariable, varValue)

Const ForWriting = 2
Const WshHide = 0
Dim objFSO, objTextFile, strFileText, aDataFile, nF
Dim bFoundSet

'open autoexec.bat
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("C:\autoexec.bat")

strFileText = objTextFile.ReadAll

'read the file into an array
aDataFile = Split(strFileText, vbCrLf)

bFoundUser = False

'search for the SET statement for the specified variable
For nF = 0 To UBound(aDataFile)
    'if it's found, then update line
    If InStr(1, aDataFile(nF), "set " & strVariable, vbTextCompare) >0 Then
        bFoundSet = True
        aDataFile(nF) = "SET " & strVariable & "=" & varValue 
        Exit For
    End If
Next

objTextFile.Close

'open autoexec.bat for writing
Set objTextFile = objFSO.OpenTextFile("C:\autoexec.bat", ForWriting)
strFileText = Join(aDataFile, vbCrLf)

'write back contents of file
objTextFile.Write strFileText
'if set statement not found, then 
If Not bFoundSet Then
objTextFile.WriteLine
    objTextFile.WriteLine "SET " & strVariable & "=" & varValue
End If

objTextFile.Close

Set objShell = CreateObject("WScript.Shell")
'run Winset. This assumes it is in the path
objShell.Run "winset " & strVariable & "=" & varValue, WshHide, True

End Sub
