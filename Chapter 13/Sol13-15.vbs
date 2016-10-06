Const adSaveCreateOverWrite = 2
Const adModeRead = 1
Const adTypeBinary = 1
Dim objRec, objStream
Set objStream = CreateObject("ADODB.Stream")

objStream.Open "URL=http://www.microsoft.com/library/toolbar/images/mslogo.gif" _ 
              , adModeRead 
'set stream type to Binary and save file
objStream.Type = adTypeBinary
objStream.SaveToFile "e:\data\mslogo.gif", adSaveCreateOverWrite
objStream.Close

