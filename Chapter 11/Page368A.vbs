Set objIE = CreateObject("ENTWSH.HTMLGen")

  objIE.StartDOC "Hello World", True
  objIE.WriteLine "<b><center>Hello World</center></b>"
  objIE.WritePara "The quick brown dog. etc.. etc.."
  objIE.EndDOC
