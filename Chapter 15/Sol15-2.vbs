'list all Web sites on the Thor Web server
Dim objWebService, objWebSite

Set objWebService = GetObject("IIS://Thor/W3SVC")

'loop through each site, find available site #
For Each objWebSite In objWebService
    'check if the object is a Web site
    If objWebSite.Class = "IIsWebServer" Then
     Wscript.Echo objWebSite.Name, objWebSite.ServerComment
    End If
Next
