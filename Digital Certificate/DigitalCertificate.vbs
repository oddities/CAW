	Option Explicit
Dim objExcel, objExcel1, objShell, objWorkbook, objPDHQ, strExcelPath
Dim strPDHQ, objSheet, objSheet2, rows, col, row, words, a, x, b
Dim objbook, IE, intRow, struid, stresetup, strrole, strgroup, comments, WshShell, strrole2
Dim FSO, strCurDir, f, name, objTextFile

Set FSO =CreateObject("scripting.FileSystemObject")

For Each f in FSO.GetFolder("C:\BMS Automation\log").Files
  name = LCase(f.Name)
  If FSO.GetExtensionName(name) = "txt" Then
    Set objTextFile = FSO.OpenTextFile ("C:\BMS Automation\log\"& f.Name, 8, True)
  End If
Next

Set objShell = CreateObject("Wscript.Shell")
objShell.Run("taskkill /im iexplore.exe"), 1, TRUE

strCurDir = objShell.CurrentDirectory  'Set the directory as current directory
FSO.CopyFile "\\USHPWBMSFSP002.ONE.ADS.BMS.COM\shared02\CAA\Groups\Digital Certificate\*", "C:\BMS Automation\Digital Certificate\",True
WScript.Sleep 5000

Set objExcel1 = CreateObject("Excel.Application")
Set objbook = objExcel1.Workbooks.Open("C:\BMS Automation\Digital Certificate\data.xlsx")
Set objSheet = objExcel1.ActiveWorkbook.Worksheets(1)

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\BMS Automation\Digital Certificate\Groups.xlsx")

x=1
rows = 1
words =0
col = 1

Do Until objExcel1.Cells(rows,1).Value =  ""

 If (x<10) Then
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = 1
  objTextFile.WriteLine(Now())
  struid  = Trim(objExcel1.Cells(rows,5).Value)
  stresetup = Trim(objExcel1.Cells(rows,1).Value)
  strrole = Trim(objExcel1.Cells(rows,14).Value)
  strrole2 = Trim(objExcel1.Cells(rows,15).Value)
  objTextFile.WriteLine(struid &" "& stresetup &" "& strrole &" "& strrole2)
  IE.navigate "https://adaccounts.ads.bms.com/showUser.aspx?uid="+struid
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Opened AD Website")
  IE.Document.All.Item("tb_eSetupControl").Value =stresetup
  intRow = 1
  Do Until objExcel.Cells(intRow,1).Value =  ""

   If(objExcel.Cells(intRow,1).Value = strrole OR objExcel.Cells(intRow,1).Value = strrole2) Then
    
    strgroup =objExcel.Cells(intRow, 2).Value
    IE.Document.All.Item("tb_manualGroupName").Value =strgroup
	objTextFile.WriteLine(Now())
    objTextFile.WriteLine("Added group :"& strgroup)
    IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
    If(objExcel.Cells(intRow, 3).Value <> "") Then
     Set XE = CreateObject 
     XE.Visible = 1
     XE.navigate "http://mygroups.bms.com/groups/SG-PKISMARTCARD-USERS"
	 objTextFile.WriteLine(Now())
     objTextFile.WriteLine("Opened Myshares website")
    End If
    If(objExcel.Cells(intRow, 4).Value <> "") Then
     
     comments = objExcel.Cells(intRow, 4).Value
     Set WshShell = WScript.CreateObject("WScript.Shell") 
     WshShell.Run "cmd.exe /c echo " & comments & " | clip", 0, TRUE
     b=Msgbox("Press Ctrl + V in Provisioner notes in eSetup "+stresetup,0," Provisioner notes comments")
	 objTextFile.WriteLine(Now())
     objTextFile.WriteLine("Comment added to clipboard to put in provisioner notes")
    End If
   End If
   intRow = intRow + 1
   Do While (IE.Busy)    
     WScript.Sleep 100    
   Loop
  Loop
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  Else
  objShell.Run("taskkill /im iexplore.exe"), 1, TRUE
  x=0
 End If
 x=x+1
 rows = rows+1
 objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Moving to next eSetup")
Loop
objTextFile.WriteLine(Now())
objTextFile.WriteLine("Completed Digital Certificate Script")
objTextFile.Close
objExcel.ActiveWorkbook.Close
objExcel1.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel1.Application.Quit
objExcel.Quit
objExcel1.Quit
FSO.DeleteFile "Groups.xlsx", True 
Msgbox "completed"