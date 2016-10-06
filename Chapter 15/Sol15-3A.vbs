Const MediumProtection = 2
Dim objService, objWebSite, objVirtDir
'get a reference to the Web service on server Odin
Set objService = GetObject("IIS://odin/W3SVC")

'create a new Web site and assign it the value of 5 as its 'name'
Set objWebSite = objService.create("IIsWebServer", 5)
'assign a friendly comment to appear in the IIS MMC
objWebSite.ServerComment = "New Web site"
objWebSite.SetInfo

'bind the site to an IP address and distinguished name
objWebSite.ServerBindings = Array("192.168.1.40:80:accounting.acme.com")
objWebSite.SetInfo

'create a virtual ROOT directory for the site
Set objVirtDir = objWebSite.create("IIsWebVirtualDir", "ROOT")
objVirtDir.Accessread = True
objVirtDir.Path = "f:\inetpub\newsite" 'set the Web directory
objVirtDir.SetInfo

'create an application for the ROOT directory
objVirtDir.AppCreate MediumProtection
objVirtDir.AppFriendlyName = "Default Application"

objVirtDir.SetInfo
