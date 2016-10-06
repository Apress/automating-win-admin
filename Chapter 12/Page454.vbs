Set objSession = CreateObject("MAPI.Session")
objSession.Logon 
'loop through an 
For Each objInfoStore In objSession.Infostores
   WScript.Echo objInfoStore.Name
Next

'reference the Personal Folder Infostore by name..
Set objInfoStore = objSession.InfoStores.Item("Personal Folder")

'reference the first Infostore object in the InfoStores collection
Set objInfoStore = objSession.InfoStores.Item(1)
