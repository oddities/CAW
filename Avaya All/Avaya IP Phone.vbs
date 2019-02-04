filename = "userdet.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename)

Do Until f.AtEndOfStream
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = 1
  line = f.ReadLine
  struid  = Split(line)(1)
  stresetup = Split(line)(0)
  strmobile = Split(line)(2)
  
  IE.navigate "https://adaccounts.ads.bms.com/showUser.aspx?uid="+struid
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  IE.Document.All.Item("tb_eSetupControl").Value =stresetup
  If (strmobile = "Yes" OR strmobile = "yes" OR strmobile = "flex" OR strmobile = "Flex" OR strmobile = "FLEX") Then
   IE.Document.All.Item("tb_manualGroupName").Value = "GG_APP_AvayaIPSoftphone-52"
   IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
  Else
   comments = "Please be advised that you are not approved for the Mobile Work program. Please visit Http://mwp.bms.com for details. If you believe you may have been approved under the Flex Worker Program, please resubmit your eSetup choosing option Avaya IP Softphone (Flex Worker) so that proper approval may be obtained"
   Set WshShell = WScript.CreateObject("WScript.Shell") 
   WshShell.Run "cmd.exe /c echo " & comments & " | clip", 0, TRUE
   b=Msgbox("Press Ctrl + V in Provisioner notes in eSetup "+stresetup,0," Provisioner notes comments")
  End If
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
Loop

f.Close 