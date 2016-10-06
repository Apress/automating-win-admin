Dim objRegExp

Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Pattern = "\w+\b"
objRegExp.Global = True

Wscript.Echo _ 
   objRegExp.Replace("FRED MCTAVISH", GetRef("Proper"))

Function Proper(strMatch, nPos, strSource)
 'check if the match starts with MC
 If StrComp(Left(strMatch,2), "Mc", vbTextCompare)=0 Then
  Proper= "Mc" & Ucase(Mid(strMatch,3,1)) & Lcase(mid(strMatch,4))
 Else
  Proper= Ucase(Left(strMatch,1)) & Lcase(mid(strMatch,2))
 End If
End Function
