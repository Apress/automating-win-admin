Set objWebService = GetObject("IIS://odin/W3SVC")
'delete the Web server named 3
objWebService.delete "IIsWebServer", 3
