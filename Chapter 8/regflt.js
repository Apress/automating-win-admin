//regflt.js
//filters regular expression elements from strings
var nF, rst, strLine, strOut;
 
if(WScript.Arguments.length< 1 || WScript.Arguments.length>1)
  WScript.Quit(-1);

// get the regular expression string
  rst = WScript.Arguments.Item(0); 

   //loop until the end of the text stream has been encountered
   while(!WScript.StdIn.AtEndOfStream)  
   {     
    strLine = WScript.StdIn.ReadLine();
   
    arg = strLine.match(rst); 

    arg = strLine.match(rst); 
    if (arg) 
    {
       for(nF=1;nF<arg.length;nF++) 
            WScript.StdOut.Write(arg[nF] + (nF<arg.length-1 ? "," : ""));
        WScript.StdOut.WriteLine();
      }
   }
