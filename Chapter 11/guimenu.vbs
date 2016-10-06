'guimenu.vbs
 'build menu in IE based on command line parameter
 Option Explicit
 Dim objIE, objDoc, aMenuItems, nReturnValue
 Dim nReturn, nF, bDone

 If Wscript.Arguments.Count <> 1 Then
  Wscript.StdErr.WriteLine _
         "You must specify a list of menu items seperated by semi-colons"
  Wscript.Quit -1
 End If

 'create an instance of the IE browser
 Set objIE = Wscript.CreateObject("InternetExplorer.Application", "ie_")
  
 'get the menu items 
 aMenuItems = Split(Wscript.Arguments(0),";")
  
 'turn off all on screen elements
 objIE.AddressBar= False
 objIE.MenuBar= False
 objIE.ToolBar= 0

 objIE.Navigate "about:blank"

 'wait to load page
 While objIE.Busy : Wend
 Set objDoc = objIE.Document

'build HTML page based on menu items
For nF = 0 To Ubound(aMenuItems) 
 objDoc.Write  "<center><input type=""button"" value=""" & _ 
           aMenuItems(nF) & """ name=""" & nF & """></p></center>"

 Set objDOC.All(Cstr(nF)).onclick = GetRef("OnButton_Click")
Next

objIE.Height = 60 * nF + 35
objIE.Width = 300

objIE.Visible=True

bDone = False
 While Not bDone 
  Wscript.sleep 100
 Wend

 Wscript.Stdout.WriteLine nReturnValue
 nReturn = nReturnValue

 'check if menu button selected, exit IE
 If nReturnValue <> -1 Then objIE.Quit

 Wscript.Quit nReturn
 
'event fires when IE is exited
 Sub ie_OnQuit
  nReturnValue = -1
  bDone = True
 End Sub

'this subroutine is called when a menu button is clicked
 Sub OnButton_Click
  nReturnValue = objdoc.activeelement.name
  bDone = True
 End Sub
