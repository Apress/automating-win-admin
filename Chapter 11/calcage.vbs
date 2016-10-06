'calcage.vbs
Dim objIE, objDoc, bDone
'create an instance of the IE browser
Set objIE = WScript.CreateObject("InternetExplorer.Application","ie_")

'turn off all menus/toolbars and set window size
objIE.AddressBar= False
objIE.MenuBar= False
objIE.ToolBar= 0
objIE.Width = 400
objIE.Height = 250

'go to the page
objIE.Navigate "c:\Code Download\Chapter 11\calcage.htm"

 'wait to load page
 While objIE.Busy
 Wend
 objIE.Visible = True  'display page
 Set objDoc = objIE.Document.All

 'assign HTML form buttons to subroutines 
 Set objDOC("cmdQuit").onclick = GetRef("cmdQuit_OnQuit")
 Set objDOC("cmdCalculate").onclick = GetRef("cmdCalculate_OnClick")
 Set objDOC("txtBirthDate").OnChange = GetRef("txtBirthDate_OnChange")
 Set objDOC("txtBirthDate").OnMouseOver = GetRef("txtStatus_Change") 
 Set objDOC("txtBirthDate").OnMouseDown = GetRef("txtStatus_Change") 
 Set objDOC("txtBirthDate").OnMouseUp = GetRef("txtStatus_Change") 

  bDone = False
  While Not bDone
    WScript.Sleep 100
  Wend

Sub txtStatus_Change
 objDOC("txtBirthDate").value = Date
End Sub

'event fires when IE is exited. 
 Sub cmdQuit_OnQuit
  objIE.Quit
  bDone = True
 End Sub

 'event fires when value in birth date field is changed
 Sub txtBirthDate_OnChange
   Call CalculateAge()
 End Sub

 'event fires when value in Full Name field is changed
 Sub cmdCalculate_OnClick
   If CalculateAge() Then 
    WScript.Echo  objDOC("txtFirstName").value & " " & _
                   objDOC("txtLastName").value & " is " & _
                   objDOC("txtAge").value
   End If
 End Sub

 'validates date field and calculates age
 Function CalculateAge
  Dim strDate
  'get the birthdate entered on the form
  strDate = objDoc("txtBirthDate").value
  
  'validate date and calulate age
  If Not IsDate(strDate) Then
   MsgBox "You must enter a valid date"  
   CalculateAge = False
  Else
   objDoc("txtAge").value = DateDiff("yyyy",CDate(strDate), Date)
   CalculateAge = True
  End If
 End Function
