'create ADO record object
Set objRecord = CreateObject("ADODB.Record")
'open a Record object to the root of Acme.com
objRecord.Open "", "URL=http://www.acme.com"
'copy data.htm from www.acme.com root to backup.htm
objRecord.CopyRecord "data.htm", "URL=http://www.acme.com/backup.htm"

'move history.htm from www.acme.com root to data directory
objRecord.MoveRecord "history.htm", "http://www.acme.com/data/history.htm"
'move data.htm from www.acme.com root to dataold.htm in same directory,
'this is the same as a rename
objRecord.MoveRecord "data.htm", " dataold.htm"
