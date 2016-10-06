Dim objFileService, objSession

'get the file service object
Set objFileService = GetObject("WinNT://Odin/LanmanServer")
'filter on file shares
objFileService.Filter = Array("FileShare")

'loop through and display description of all file shares
For Each objFileShare In objFileService
 Wscript.Echo objFileShare.Name
Next
