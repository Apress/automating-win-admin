'logon.vbs
 Option Explicit

 Dim objIE, bDone, objDoc

'create an instance of the IE browser. Allow IE events to be caught
'by specifying the second parameter, ie_
 Set objIE = WScript.CreateObject("InternetExplorer.Application","ie_")

 'turn off all on screen 'clutter'
 objIE.MenuBar = False
 objIE.ToolBar = 0
 objIE.Height = 350 'resize browser form
 objIE.Width = 550

 'select the page to display
 objIE.Navigate "c:\Code Download\Chapter 11\welcome.htm" 'wait to load page
 While objIE.Busy
  WScript.Sleep 100
 Wend

 objIE.Visible = True

 bDone = False
 
 'link to the browser document
 Set objDoc = objIE.Document
 'assign the onclick event of the HTML pages cmdContinue button to the 
 'cmdContinue_OnClick sub routine in this script
 Set objDOC.All("cmdContinue").onclick = GetRef("cmdContinue_OnClick") 
 'loop until done
 While Not bDone 
   wscript.sleep 100
 Wend

 'this event is fired when IE is exited
 Sub ie_OnQuit
  bDone = True
 End Sub
 
 Sub cmdContinue_OnClick
  objIE.Quit
  bDone = True
 End Sub
