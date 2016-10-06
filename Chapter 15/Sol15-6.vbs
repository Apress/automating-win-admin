'get a reference to the root of the first Web site
Set objContainer = GetObject("IIS://odin/W3SVC/1/Root")

'create a virtual directory called AcctDir
Set objVirtDir = objContainer.Create("IIsWebVirtualDir", "AcctDir")
'set the virtual directory to point to a local folder
objVirtDir.Path = "d:\data\sites\intranet"
objVirtDir.SetInfo
