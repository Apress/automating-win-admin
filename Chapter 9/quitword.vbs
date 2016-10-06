'quitword.vbs
'Finds a running copy of Word, saves and closes all files
Dim objWord, objDoc

 'get an instance to an existing copy of Word
 Set objWord = GetObject(, "Word.Application")

 'loop through each Word document
 For Each objDoc In objWord.Documents
  'check if the path is empty - this identifies
  If objDoc.Path = "" Then
     objDoc.SaveAs objDoc.Name
     objDoc.Close
  Else
     objDoc.Close True
  End If
 Next
 
 'quit Word
 objWord.Quit

 Set objWord = Nothing
 Set objDoc = Nothing
