Dim objService, objFTPSite, objVirtDir

'get a reference to the FTP service on server Thor 
Set objService = GetObject("IIS://thor/MSFTPSVC")

'create a new FTP site and assign it the value of 6 as it's 'name'
Set objFTPSite = objService.create("IIsFTPServer", 6)
'assign a friendly comment to appear in the IIS MMC
objFTPSite.ServerComment = "Accounting FTP Site"

'bind the site to an IP address and distinguished name
objFTPSite.ServerBindings = Array("192.168.1.40:21:ftp.accounting.acme.com")
objFTPSite.SetInfo

'create a virtual ROOT directory for the site
Set objVirtDir = objFTPSite.create("IIsFTPVirtualDir", "ROOT")
objVirtDir.Accessread = True
objVirtDir.Path = "f:\inetpub\ftp" 'set the FTP directory
objVirtDir.SetInfo
