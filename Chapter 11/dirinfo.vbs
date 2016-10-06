'dirinfo.vbs
 Dim objIE, objFSO 
 Dim objFolder, objSrcFolder 

 Set objIE = CreateObject("ENTWSH.HTMLGen")
 Set objFSO = CreateObject("Scripting.FileSystemObject")

 On Error Resume Next
  
 objIE.StartDOC "Folder Size", True
 objIE.WriteLine "<c><b><h2>User Folder Size, over 5 megs</c></b></h2>"
 objIE.StartTable Array(100, 300), "0" 
 objIE.WriteRow Array("<b>Folder Name","<b>Size"), "bgcolor=""#FFFF00"""
    
  For Each objFolder In objFSO.GetFolder("c:\").SubFolders
    If objFolder.Size > 5000000 Then
      objIE.WriteRow Array(objFolder.Name, objFolder.Size), ""
    End If
  Next
  
  objIE.EndTable
  objIE.EndDOC
  
  strHTML = objIE.HTML
