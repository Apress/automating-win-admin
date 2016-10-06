'listprop.vbs
'lists schema properties for specified object
Dim objClass,varAttrib, aval, objObject, strLine, objSchema
On Error Resume Next

 If Wscript.Arguments.Count <> 1 Then 
   ShowUsage
   Wscript.Quit
 End If 

 'get the object
 Set objObject = GetObject(Wscript.Arguments (0))

 If Err Then
  Wscript.Echo "Unable to get object " & Wscript.Arguments (0)
  Wscript.Quit
 End If
 
 Set objClass = GetObject(objObject.Schema)

 Set objSchema = GetObject(objClass.Parent)

 Wscript.Echo "Mandatory Attributes: "

 For Each varAttrib In objClass.MandatoryProperties
  strLine = "   " & varAttrib
     
   'if array then display all items in array
     If IsArray(objObject.Get(varAttrib)) Then
        If Not Err Then
         For Each aval In objObject.Get(varAttrib)
            varAttrib = varAttrib & "," & aval
         Next
        End If
    Else
        'if property is object attempt to determine what type and
        'display value
        If IsObject(objObject.Get(varAttrib)) Then
      
            Set objProp = objSchema.GetObject("Property", varAttrib)
            'if object is INT8 then convert
            If objProp.Syntax = "INTEGER8" Then
                strLine = strLine & " " & ConverINT8(objObject.Get(varAttrib))
            Else
                strLine = strLine & " Object type:" & objProp.Syntax
            End If
            
        Else
            strLine = strLine & " " & objObject.Get(varAttrib)
        End If
    End If
    If Err Then strLine = strLine & " No value"

    Err.Clear
    Wscript.Echo strLine
 Next

 Debug.Print "Optional Attributes: "
 For Each varAttrib In objClass.OptionalProperties
   strLine = "   " & varAttrib
   objObject.GetInfoEx Array(varAttrib), 0
   'check if object is an array
    If IsArray(objObject.Get(varAttrib)) Then
        If Not Err Then
         For Each aval In objObject.Get(varAttrib)
            strLine = strLine & "," & aval
         Next
        End If
    Else
        'if property is object attempt to determine what type and
        'display value
        If IsObject(objObject.Get(varAttrib)) Then
            Set objProp = objSchema.GetObject("Property", varAttrib)
            Wscript.Echo "opbject"
            'if object is INT8 then convert
            If objProp.Syntax = "INTEGER8" Then
                strLine = strLine & " " & ConvertINT8(objObject.Get(varAttrib))		
            Else
                strLine = strLine & " Object type:" & objProp.Syntax
            End If
            
        Else
            strLine = strLine & " " & objObject.Get(varAttrib)
        End If
    End If
        If Err Then strLine = strLine & " No value"
    Err.Clear
  Wscript.Echo strLine
 Next

Sub ShowUsage
  WScript.Echo "listprop list properties for specified object" _
    & vbCrLf &  "Syntax:" &  vbCrLf & _
    "listprop.vbs objectpath" &  vbCrLf & _
     "objectpath Path to " & vbCrLf & _
     "Example: List details about computer Odin" & vbCrLf & _
     "listprop WinNT://Odin,computer"
End Sub 

'convert INT8 value to number
Function ConvertINT8(objTime)
Dim nLow

'get the lower 32 bits
nLow = objTime.LowPart
'if negative value add 2^32 to low value
If nLow < 0 Then
    nLow = nLow + 2 ^ 32
End If

ConvertINT8 = CStr(Abs((objTime.HighPart * 2 ^ 32) + nLow))

End Function