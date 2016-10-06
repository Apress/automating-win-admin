Dim objSecurity, objSD, strSource, strDest, objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSecurity = CreateObject("ADsSecurity")
strSource = "d:\data\report.doc"
strDest = "d:\backup\report.doc"

'copy the file to the destination
objFSO.CopyFile strSource, strDest
' get the security descriptor for the file
 Set objSD = _
       objSecurity.GetSecurityDescriptor("FILE://" & strSource)

 'copy the security descriptor from the original file to the copied file
 objSecurity.SetSecurityDescriptor objSD, "FILE://" & strDest
