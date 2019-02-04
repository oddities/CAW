filename = "userdet.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename)

Do Until f.AtEndOfStream
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = 1
  line = f.ReadLine
  struid  = Split(line)(1)
  stresetup = Split(line)(0)
  
  IE.navigate "https://adaccounts.ads.bms.com/showUser.aspx?uid="+struid
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  IE.Document.All.Item("tb_eSetupControl").Value =stresetup
  IE.Document.All.Item("tb_manualGroupName").Value = "GG_IMSS_METAFRAME_BIOCON_ELN_DESKTOP_USERS"
  IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
Loop

f.Close 