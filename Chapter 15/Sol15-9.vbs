Set objSite = GetObject("IIS://thor/W3SVC/4")
'bind two domain names to the site
objSite.ServerBindings = Array("192.168.1.40:80:sales.acme.com", _
                          "192.168.1.40:80:marketing.acme.com")

objSite.SetInfo
