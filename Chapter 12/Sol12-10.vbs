Dim objWebems, objWeb, objWebem, sMsg, st, objMail

Set objWebem = GetObject("winmgmts:\\")
Set objWebems = objWebem.ExecQuery("SELECT TimeWritten, EventIdentifier " & _
    ", Message  FROM Win32_NTLogEvent WHERE EventIdentifier=529 " & _
    "AND SourceName='Security' AND TimeWritten > '" & _
    Convert2DMTFDate(Date - 1, "630") & "' ")

sMsg=""
For Each objWeb In objWebems
    st = objWeb.Properties_("TimeWritten")
   
    st = Mid(st, 5, 2) & "/" & Mid(st, 7, 2) & "/" & Mid(st, 1, 4) & _
        "  " & Mid(st, 9, 2) & ":" & Mid(st, 11, 2) & ":" & Mid(st, 13, 2)
   
sMsg = sMsg & "Time Logged " & st & vbCrLf
sMsg = sMsg & Replace(objWeb.Properties_("Message"),_ 
        vbCrLf & vbCrLf, vbCrLf)
Next

If Not sMsg="" Then
  Set objMail = CreateObject("CDO.Send")
 objMail.Profile = "My Profile"  
 objMail.NewMessage  
 objMail.AddRecipient ("SMTP:administrator@abc.com")  
 objMail.Message = sMsg
 objMail.Subject = "Logon failures detected"  
 objMail.Send 
 objMail.Logoff 
End If

'convert date to DMTF date required by WMI
Function Convert2DMTFDate(dDate, sTimeZone)
Dim sTemp
sTemp = Year(Now) & Pad(Month(dDate), 2, "0") & Pad(Day(dDate), 2, "0")
sTemp = sTemp & Pad(Hour(dDate), 2, "0") & Pad(Minute(dDate), 2, "0")
sTemp = sTemp & "00.000000+" & sTimeZone
Convert2DMTFDate = sTemp
End Function

'pad a string
Function Pad(sPadString, nWidth, sPadChar)
    If Len(sPadString) < nWidth Then
        Pad = String(nWidth - Len(sPadString), sPadChar) & sPadString
    Else
        Pad = sPadString
    End If
End Function
