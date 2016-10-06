'GetDataType
'Returns ADO data type.
'Parameter
'nType  ADO Data type value 
'Returns string value representing data type name
Function GetDataType(nType)
 Dim strRet

Select Case nType
 Case 20
  strRet = "adBigInt"
 Case 128
  strRet = "adBinary"
 Case 11
  strRet = "adBoolean"
 Case 8
  strRet = "adBSTR"
 Case 136
  strRet = "adChapter"
 Case 129
  strRet = "adChar"
 Case 6
  strRet = "adCurrency"
 Case 7
  strRet = "adDate"
 Case 133
  strRet = "adDBDate"
 Case 137
  strRet = "adDBFileTime"
 Case 134
  strRet = "adDBTime"
 Case 135
  strRet = "adDBTimeStamp"
 Case 14
  strRet = "adDecimal"
 Case 5
  strRet = "adDouble"
 Case 0
  strRet = "adEmpty"
 Case 10
  strRet = "adError"
 Case 64
  strRet = "adFileTime"
 Case 72
  strRet = "adGUID"
 Case 9
  strRet = "adIDispatch"
 Case 3
  strRet = "adInteger"
 Case 13
  strRet = "adIUnknown"
 Case 205
  strRet = "adLongVarBinary"
 Case 201
  strRet = "adLongVarChar"
 Case 203
  strRet = "adLongVarWChar"
 Case 131
  strRet = "adNumeric"
 Case 138
  strRet = "adPropVariant"
 Case 4
  strRet = "adSingle"
 Case 2
  strRet = "adSmallInt"
 Case 16
  strRet = "adTinyInt"
 Case 21
  strRet = "adUnsignedBigInt"
 Case 19
  strRet = "adUnsignedInt"
 Case 18
  strRet = "adUnsignedSmallInt"
 Case 17
  strRet = "adUnsignedTinyInt"
 Case 132
  strRet = "adUserDefined"
 Case 204
  strRet = "adVarBinary"
 Case 200
  strRet = "adVarChar"
 Case 12
  strRet = "adVariant"
 Case 139
  strRet = "adVarNumeric"
 Case 202
  strRet = "adVarWChar"
 Case 130
  strRet = "adWChar"
 End Select

 GetDataType = strRet
End Function
