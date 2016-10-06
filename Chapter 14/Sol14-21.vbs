'get the file service object
Set objFileService = GetObject("WinNT://Odin/LanmanServer")

Set objFileShare = objFileService.create("FileShare", "AcctData")
objFileShare.Path = "d:\data\accounting"
objFileShare.Description = "Accounting Data" 
objFileShare.MaxUserCount = -1 'set unlimited users
objFileShare.SetInfo
