Const adReadLine = -2
Const adTypeText = 2
'read the contents of a Web file
Set objStream = CreateObject("ADODB.Stream")
'open to log.txt file on www.acme.com site
objStream.Open "URL=http://www.acme.com/log.txt" 
'set type to text and character set to Ascii
objStream.Type = adTypeText
objStream.charset = "ascii"

'read contents of file and output 
Do While Not objStream.EOS
 'read the next line of text
  strLine = objStream.ReadText(adReadLine)
  Wscript.Echo strLine
Loop
objStream.Close
