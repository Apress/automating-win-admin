Dim objServices, nResult, objFile

'connect to remote computer Odin
Set objServices = GetObject("winmgmts:{impersonationLevel=impersonate}!\\Odin")

'get a reference to file reports.doc
Set objFile = _
          objServices.Get("CIM_DataFile.Name='d:\\data\\reports.doc'")

'copy file to backup reports directory
nResult = objFile.Copy("d:\\backup\\reports\reports.doc")
