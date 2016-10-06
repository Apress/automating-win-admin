Dim objIE 
Set objIE = CreateObject("InternetExplorer.Application")
objIE.Navigate "about:blank"

objIE.Visible = True
objIE.Document.Write "<b>hello world</b>"
