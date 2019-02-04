Option Explicit
Dim objExcel, objExcel1, objShell, objWorkbook, objPDHQ, strExcelPath
Dim strPDHQ, objSheet, objSheet2, rows, col, row, words, a, x, b, I, c
Dim objbook, IE, intRow, struid, strbmsid, stresetup, strrole, strrole1, strcountry, comments, WshShell, element, htmlelement
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
FSO.CopyFile "\\USHPWBMSFSP002.ONE.ADS.BMS.COM\shared02\CAA\Groups\DIMS\*", "C:\BMS Automation\DIMS\",True
WScript.Sleep 5000

Set objExcel1 = CreateObject("Excel.Application")
Set objbook = objExcel1.Workbooks.Open("C:\BMS Automation\DIMS\data.xlsx")
Set objSheet = objExcel1.ActiveWorkbook.Worksheets(1)

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\BMS Automation\DIMS\Groups.xlsx")

x=1
rows = 1
words =0
col = 1

Do Until objExcel1.Cells(rows,1).Value =  ""
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = 1
  objTextFile.WriteLine(Now())  
  strrole = Trim(objExcel1.Cells(rows,13).Value)
  struid  = Trim(objExcel1.Cells(rows,5).Value)
  stresetup = Trim(objExcel1.Cells(rows,1).Value)
  strcountry = Trim(objExcel1.Cells(rows,12).Value)
  objTextFile.WriteLine(struid &" "& stresetup &" "& strrole)
  
  IE.navigate "https://adaccounts.ads.bms.com/showUser.aspx?uid="+struid
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Opened AD Website")
  
  IE.Document.All.Item("tb_eSetupControl").Value =stresetup
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Added eSetup number :"& stresetup)
  intRow = 1
  Do Until objExcel.Cells(intRow,1).Value =  ""
   If (objExcel.Cells(intRow,1).Value = strrole) Then
    If (objExcel.Cells(intRow,2).Value = strcountry) Then
     IE.Document.All.Item("tb_manualGroupName").Value = objExcel.Cells(intRow, 3).Value
	 objTextFile.WriteLine(Now())
     objTextFile.WriteLine("Added Group : "& objExcel.Cells(intRow, 3).Value)
     IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
    ElseIf(objExcel.Cells(intRow,2).Value = "") Then
     IE.Document.All.Item("tb_manualGroupName").Value = objExcel.Cells(intRow, 3).Value
	 objTextFile.WriteLine(Now())
     objTextFile.WriteLine("Added Group : "& objExcel.Cells(intRow, 3).Value)
     IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
    End If
   End If
   
   intRow = intRow+1
  Loop
  
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  rows = rows + 1
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Moving to next eSetup")
Loop

objTextFile.WriteLine(Now())
objTextFile.WriteLine("Completed DIMS Script")
objTextFile.Close
objExcel.ActiveWorkbook.Close
objExcel1.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel1.Application.Quit
objExcel.Quit
objExcel1.Quit
FSO.DeleteFile "Groups.xlsx", True 
Msgbox "completed"
 