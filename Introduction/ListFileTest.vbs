'create an instance of the ListFiles component 
Dim objListFiles 
Set objListFiles = Wscript.CreateObject ("ENTWSH.ListFiles","lf_")

objListFiles.Path = "d:\data" 'set the path to search
'call the search method
Call objListFiles.Search

Set objListFiles = Nothing
'FoundFile event, this is called for each file found in the search
'directory
Sub lf_FoundFile(strPath)
  Wscript.Echo strPath
End Sub
