Set objRegExp = New RegExp
' match "digits slash digits slash digits"
objRegExp.Pattern = "(\d+)\/(\d+)\/(\d+)"
objRegExp.IgnoreCase = True 

Set objMatches = objRegExp.Execute("5/13/2000")

'list all sub-expressions 
For nF = 0 To objMatches(0).SubMatches.Count - 1
  Wscript.Echo objMatches(0).SubMatches(nF)
Next
