'create an instance of the IE browser
Set objIE = CreateObject("InternetExplorer.Application")

'build a page containing a text field prompting for a date
objIE.Navigate "about:blank"

'wait to load page
While objIE.Busy : Wend


objIE.Document.body.innerHTML = _ 
             "<html>Date <input type=""text"" name=""txtDate"" size=""10""</html>"

objIE.Visible = True  'display page

Set objDoc = objIE.Document.All
'set the value for the field txtDate to today's date
objDoc("txtDate").Value = Date
