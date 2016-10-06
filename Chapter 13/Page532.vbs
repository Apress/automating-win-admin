'copy a local file to a Web server
Option Explicit
Const adTypeBinary = 1
Const adCreateNonCollection = 0
Const adCreateOverwrite = &H4000000 
Const adModeReadWrite = 3
Const adOpenStreamFromRecord=4

Dim objRst, objConn, objRecord, objRec, objStream 
Set objStream = CreateObject("ADODB.Stream")
Set objRecord = CreateObject("ADODB.Record")

'create a new data.dat file, overwriting the any file with the same name
objRecord.Open "data.zip", "URL=http://www.acme.com/" _
               , adModeReadWrite, adCreateNonCollection + adCreateOverwrite

'open the Stream object using the objRecord object
objStream.Open objRecord, adModeReadWrite, adOpenStreamFromRecord
objStream.Type = adTypeBinary \
'load local file in Stream object
objStream.LoadFromFile "d:\data\data.dat"
objStream.Close
objRecord.Close
