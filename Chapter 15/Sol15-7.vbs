'get a reference to the file data.txt in Root directory of site 3 on Thor
Set objFile = GetObject("IIS://thor/W3SVC/3/Root/data.txt")
objFile.AccessWrite = False
objFile.SetInfo
