'opt.vbs
'performs operation based on selected menu item
Dim strOption
strOption = WScript.Stdin.ReadLine

Select Case strOption
 Case "0"
  Wscript.Echo "Option 0 was selected..."
 Case "1"
  Wscript.Echo "Option 1 was selected..."
End Select
