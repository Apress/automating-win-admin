'pingit.vbs
Dim strLine, nCount, objTextStream, strComp, bBad
nCount = 0
'loop until the end of the text stream has been encountered
Do While Not WScript.StdIn.AtEndOfStream
  strLine = Wscript.StdIn.ReadLine
  'check if bad IP address encountered
  If Left(strLine, 14) = "Bad IP address" Then
    bBad = True
    Exit Do
  End If

  If Left(strLine, 10) = "Reply from" Then
    nCount = nCount + 1
  End If
  If Left(strLine, 7) = "Pinging" Then
    strComp = Mid(strLine, 9, Instr(strLine, "]") - 8)
  End If
Loop

'check if bad IP address encountered
If bBad Then 
  WScript.Echo "Bad IP address:" & Mid(strLine,15)
Else
  Wscript.Echo nCount & " replies received from " & strComp
End If
