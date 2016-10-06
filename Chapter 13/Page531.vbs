<?xml version="1.0" ?>
<job>
<reference object="ADODB.Stream"/> 
<!--comment
Script:openstream.wsf
Creates and opens a Stream object
-->
 <script language="VBScript">
  Dim objStream
  Set objStream = CreateObject("ADODB.Stream")
  objStream.Open "URL=http://thor/visitors.txt", adModeReadWrite
  objStream.Close
  </script>
</job>
