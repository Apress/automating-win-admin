Dim objService, objWMIObject, nResult

'get a reference to a WMI service
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}")

'get a reference to the directory D:\Data\Reports and compress it
Set objWMIObject = _
             objService.Get("Win32_Directory.Name=""D:\\Data\\Reports""")
nResult = objWMIObject.Compress

'get a reference to the file D:\Data\data.mdb and compress it
Set objWMIObject = _
             objService.Get("CIM_DataFile.Name=""D:\\Data\\data.mdb""")
nResult = objWMIObject.Compress
