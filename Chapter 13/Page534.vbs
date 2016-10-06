<?xml version="1.0" ?>
<job>
<reference object="ADODB.Stream"/> 
<!--comment
Script:write2stream.wsf
Writes to a text stream
-->
 <script language="VBScript">
  Option Explicit
  Dim objStream
  Set objStream = CreateObject("ADODB.Stream")
  'open Stream object and set character type to text and ASCII
  objStream.Open "URL=http://thor/log.txt", adModeReadWrite
  objStream.Type = adTypeText
  objStream.charset = "ascii"
  'set position to end of stream
  objStream.Position = objStream.Size
  'write to Stream and close it
  objStream.WriteText "Operation successful", adWriteLine
  objStream.Close
  </script>
</job>
