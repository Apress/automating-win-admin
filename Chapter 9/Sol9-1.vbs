On Error Resume Next
 'attempt to get an existing running copy of Word
 Set objWord = GetObject(, "Word.Application")
 'if error occurred, then couldn't find Word, create new instance
 If Err Then
    Err.Clear
    Set objWord = CreateObject("Word.Application")
 End If
 objWord.Documents.Add
 objWord.Selection.TypeText "Hello World!"
