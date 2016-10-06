'ISPDC.VBS
'checks if computer is a Primary Domain Controller (PDC)
Dim objWMIObject, objServices, objNetwork, nF, bFound, objRole

Set objNetwork = CreateObject("WScript.Network")

'get an instance of the Win32_ComputerSystem for the local computer
Set objWMIObject = _
             GetObject("winmgmts:{impersonationLevel=impersonate}" & _
             "!Win32_ComputerSystem='" & objNetwork.ComputerName & "'")

bFound = False

'check if Roles property is Empty array
If Not IsNull(objWMIObject.roles) Then
  'loop through roles array and check if it contains any occurrence
  'of Primary_Domain_Controller
   For Each objRole In objWMIObject.Roles
    If objRole = "Primary_Domain_Controller" Then
        bFound = True
        Exit For
    End If
  Next
End If

'if PDC then return -1, otherwise 0
If bFound Then WScript.Quit -1
WScript.Quit 0
