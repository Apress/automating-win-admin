'getimage.vbs
Const adSaveCreateOverWrite = 2
Const adTypeBinary = 1
Function GetImage(strPath, strDest)
  Dim objXMLHTTP, nF, arr, objFSO, objFile
  Dim objRec, objStream

  'create XMLHTTP component
  Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")

  'get the image specified by strPath
  objXMLHTTP.Open "GET", strPath, False
  objXMLHTTP.Send

  'check if retrieval was successful
  If objXMLHTTP.statusText = "OK" Then
    'create binary stream to write image output
    Set objStream = CreateObject("ADODB.Stream")
     objStream.Type = adTypeBinary
     objStream.Open
     objStream.Write objXMLHTTP.ResponseBody
     objStream.SavetoFile strDest, adSaveCreateOverwrite
     objStream.Close
     GetImage = "OK"
  Else
     GetImage = objXMLHTTP.statusText
  End If
End Function
