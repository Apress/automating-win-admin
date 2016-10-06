VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'BAConv Convert class
'Wrapper object that provides variant to byte and
'byte to variant array conversions
Private strError As String

'ByteToVariant
'Converts an array of byte values to an array of variants
'Parameters:
'aVal   Array of byte
'Returns:
'Array of variants
Public Function ByteToVariant(aVal() As Byte) As Variant
Dim arrB() As Variant
    On Error GoTo Byte2VariantError
         'check if array passed contains values
90    If UBound(aVal) <= 0 Then Exit Function
    
100    ReDim arrB(LBound(aVal) To UBound(aVal))
105    For nf = LBound(aVal) To UBound(aVal)
110        arrB(nf) = aVal(nf)
       Next
    
120    ByteToVariant = arrB
Exit Function
Byte2VariantError:
strError = "Error:" & Err.Description & " Error#:" & Err.Number & vbCrLf & _
           " Line# " & Erl
End Function

'VariantToByte
'Converts an array of variant values to byte values
'Parameters:
'aVal  Array of variants
'Returns:
'Array of Byte values
Public Function VariantToByte(aVal)
Dim arrB() As Byte
    On Error GoTo VariantToByteError
   
90    If UBound(aVal) <= 0 Then Exit Function
    
100    ReDim arrB(LBound(aVal) To UBound(aVal))
105    For nf = LBound(aVal) To UBound(aVal)
110        arrB(nf) = CByte(aVal(nf))
       Next
    
120    VariantToByte = arrB
Exit Function
VariantToByteError:
strError = "Error:" & Err.Description & " Error#:" & Err.Number & vbCrLf & _
           " Line# " & Erl
End Function

'ErrorDescription
'Returns error number, descript and line#
Public Property Get ErrorDescription() As String
    ErrorDescription = strError
End Property
