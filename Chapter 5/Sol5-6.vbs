Dim objFSO
'create an instance of an FSO object
Set objFSO = CreateObject("Scripting.FileSystemObject")

'check if the specified file exists
    If objFSO.FileExists("d:\data\report.doc") Then
        WScript.Echo "File exists"
    End If 
