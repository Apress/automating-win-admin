'compact.vbs
'compacts a database
Dim objJRO, objFSO,strSource, strTemp
strSource = "d:\test.mdb"
strTemp = "d:\temp.mdb"
Set objFSO = CreateObject("Scripting.FileSystemObject")

'check if temporary file exists from previous operation, if so delete it
If objFSO.FileExists(strTemp) Then objFSO.DeleteFile strTemp

'create Jet Replication Object..
Set objJRO = CreateObject("JRO.JetEngine")
On Error Resume Next
'compact data
objJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & strSource & ";Jet OLEDB:Engine Type=4", _ 
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTemp & _
         ";Jet OLEDB:Engine Type=4"

'check if error occurred...
If Not Err Then
    'double check the file was compacted
    If objFSO.FileExists(strTemp) Then
        'copy compacted temporary file to original
        objFSO.CopyFile strTemp, strSource, True
        objFSO.DeleteFile strTemp
    End If
Else
End If
