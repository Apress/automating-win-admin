Dim objSite, nF, aBindings
'get a reference to the fourth Web site on server Thor
Set objSite = GetObject("IIS://thor/W3SVC/4")
aBindings = objSite.ServerBindings
ReDim Preserve aBindings(UBound(aBindings) + 1)
aBindings(UBound(aBindings)) = "192.168.1.40:80:marketing.acme.com"
objSite.ServerBindings = aBindings
objSite.SetInfo
