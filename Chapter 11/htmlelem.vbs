'htmlelem.vbs
 'lists all elements in document specified by command line argument
 Dim objIE, objDoc

 If WScript.Arguments.Count <> 1 Then
   WScript.StdErr.WriteLine "You must specify a URL to process"
   WScript.Quit
 End If

 'create an instance of the IE browser
 Set objIE = CreateObject("InternetExplorer.Application")
 
 'go to the page
 objIE.Navigate Wscript.Arguments(0)

 'wait to load page
 While objIE.Busy : Wend

 Set objDoc = objIE.Document.All

 'loop through all elements in the HTML document
 For Each objItem In objDOC

 'check if the element is an input element (input box, check box etc.)
  If TypeName(objItem) = "HTMLInputElement" Then 
    Wscript.Echo objItem.value, objItem.name, objItem.type
  Else
    Wscript.Echo TypeName(objItem) 'just output the HTML object type 
  End If
 Next 
