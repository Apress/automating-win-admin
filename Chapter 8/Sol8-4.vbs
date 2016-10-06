Dim objRegExp, strName

Set objRegExp = New RegExp
 objRegExp.IgnoreCase = True
 objRegExp.Pattern = "^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$" 

 Wscript.Echo objRegExp.Replace("5/3/2000","$2/$1/$3")
