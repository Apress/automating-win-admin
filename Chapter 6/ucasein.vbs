'ucasein.vbs
'converts standard Input stream to upper case
'and redirects to stdout Dim strText
strText = WScript.StdIn.ReadAll
WScript.StdOut.Write Ucase(strText)

