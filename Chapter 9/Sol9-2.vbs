'get a reference to a document
Set objDoc = GetObject("d:\data\word\report.doc")

On Error Resume Next
For Each objProp In objDoc.BuiltinDocumentProperties
  Wscript.Echo objProp.Name, objProp.Value
Next
