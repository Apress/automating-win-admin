Function ExtractCSV(strCSV)

 Dim objRegExp, objMatch, aRet, objMatches, nF
 Set objRegExp = New RegExp   
' matches digits, digits followed by a point, digits followed by a 
' point and then more digits, a point followed by digits, or anything
' enclosed by double-quotes
  objRegExp.Pattern = "\d+\.?\d*|\.\d+|\x22[^""]+\x22" 
' matches digits, digits followed by a point, digits followed by a 
' point and then more digits, a point followed by digits, or anything
' enclosed by double-quotes
objRegExp.IgnoreCase = True   
objRegExp.Global = True    
  Set objMatches = objRegExp.Execute(strCSV)     
  If objMatches.Count > 0 Then
  
   ReDim aRet(objMatches.Count)
   For nF = 0 To objMatches.Count - 1  ' iterate Matches collection.
    Set objMatch = objMatches.item(nF)
   ' check if the string is surrounded by quotes, if so remove them
    If Left(objMatch.Value, 1) = """" And _ 
         Right(objMatch.Value, 1) = """" Then
     aRet(nF) = Mid(objMatch.Value, 2, Len(objMatch.Value) - 2)
    Else
     aRet(nF) = objMatch.Value
    End If
   Next
    
   ExtractCSV = aRet
  Else
   ExtractCSV = Empty
  End If
  
End Function
