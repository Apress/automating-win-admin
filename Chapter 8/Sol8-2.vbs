Set objRegExp = New RegExp
'set pattern to extract all numeric values from string
objRegExp.Pattern = "\d+\.?\d*|\.\d+" objRegExp.IgnoreCase = True
objRegExp.Global = True
Set objMatches = objRegExp.Execute("111.13,1232,ABC,444,55")

For Each objMatch In objMatches
  Wscript.Echo "Found match:" & objMatch.Value & " at position " & _
                objMatch.FirstIndex
Next
