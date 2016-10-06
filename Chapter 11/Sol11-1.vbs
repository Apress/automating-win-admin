  Dim objXMLHTTP
  Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")

  'get the data.htm page from www.acme.com
  objXMLHTTP.Open "GET", "http://www.acme.com/data.htm", False
  objXMLHTTP.send

  'check if retrieval was successful
  If objXMLHTTP.statusText = "OK" Then
      WScript.Echo objXMLHTTP.responseText
  Else
      WScript.Echo "Error getting page:" & objXMLHTTP.statusText
  End If
