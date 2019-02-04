Option Explicit
Dim objExcel, objExcel1, objShell, objWorkbook, strExcelPath
Dim objSheet, objSheet2, rows, col, row, words, a, x, b, c, I
Dim objbook, IE, intRow, struid, strbmsid, stresetup, strrole, strcountry, strtype, comments, WshShell, element, htmlelement
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
FSO.CopyFile "\\USHPWBMSFSP002.ONE.ADS.BMS.COM\shared02\CAA\Groups\DeltaV\*", strCurDir,True
WScript.Sleep 5000

Set objExcel = CreateObject("Excel.Application")
Set objbook = objExcel.Workbooks.Open("C:\BMS Automation\DeltaV\data.xlsx")
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

Set objExcel1 = CreateObject("Excel.Application")
Set objWorkbook = objExcel1.Workbooks.Open("C:\BMS Automation\DeltaV\Groups.xlsx")

Set I = WScript.CreateObject("WScript.Shell")

x=1
rows = 1
words =0
col = 1

Do Until objExcel.Cells(rows,1).Value =  ""
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = 1
  objTextFile.WriteLine(Now())
  struid  = Trim(objExcel.Cells(rows,5).Value)
  stresetup = Trim(objExcel.Cells(rows,1).Value)
  strbmsid = Trim(objExcel.Cells(rows,4).Value)
  strrole = Trim(objExcel.Cells(rows,14).Value)
  strtype = Trim(objExcel.Cells(rows,15).Value)
  objTextFile.WriteLine(struid &" "& stresetup &" "& strrole)
  IE.navigate "https://adaccounts.ads.bms.com/showUser.aspx?uid="+struid
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Opened AD Website")
  IE.Document.All.Item("tb_eSetupControl").Value =stresetup

  intRow = 1
  Do Until objExcel1.Cells(intRow,1).Value =  ""
   
   If(strrole = objExcel1.Cells(intRow,1).Value) Then

    If(strtype = objExcel1.Cells(intRow,2).Value) Then
     objTextFile.WriteLine(Now())
     objTextFile.WriteLine("Added Group : "& objExcel1.Cells(intRow,3).Value)   
	 IE.Document.All.Item("tb_manualGroupName").Value = objExcel1.Cells(intRow,3).Value
    End If 
   End If
   intRow = intRow + 1

  Loop
  IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Saved Changes")
  rows = rows+1
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Moving to next eSetup")
Loop
objTextFile.WriteLine(Now())
objTextFile.WriteLine("Completed Script")
objTextFile.Close
objExcel.ActiveWorkbook.Close
objExcel1.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel1.Application.Quit
objExcel.Quit
objExcel1.Quit
FSO.DeleteFile "Groups.xlsx", True 
Msgbox "completed"