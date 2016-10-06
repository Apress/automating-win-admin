Set objIE = CreateObject("ENTWSH.HTMLGen")

  objIE.StartDOC "Phone List", True
  
 'start a table with 3 columns and border width 0.
  objIE.StartTable (Array(100, 300)), "0"
  objIE.WriteRow (Array("<b>Folder Name", "<b>Size")), "bgcolor=""#FFFF00"""
  objIE.EndDOC

  objIE.WriteLine "<b><center>Phone List</center></b>"
  objIE.WritePara "The quick brown dog.. etc.. etc.."
  objIE.EndDOC
