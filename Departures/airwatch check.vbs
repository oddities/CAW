Option Explicit
Dim objExcel, objPDHQ, strExcelPath, strPDHQ, objSheet, objSheet2, rows, col, row, words, a, c, x, IE, objShell, objWorkbook
Dim struid, strspec, strsite, stresetup, I, t

Dim FSO, f, name, objTextFile

Set FSO =CreateObject("scripting.FileSystemObject")

For Each f in FSO.GetFolder("C:\BMS Automation\log").Files
  name = LCase(f.Name)
  If FSO.GetExtensionName(name) = "txt" Then
    Set objTextFile = FSO.OpenTextFile ("C:\BMS Automation\log\"& f.Name, 8, True)
  End If
Next

Set objShell = CreateObject("Wscript.Shell")
objShell.Run("taskkill /im iexplore.exe"), 1, TRUE

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\BMS Automation\Departures\data.xlsx")
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

rows = 1
words =0
col = 1
row = 1

Do Until objExcel.Cells(rows,1).Value =  ""
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = 1
  objTextFile.WriteLine(Now())
  struid = Trim(objExcel.Cells(rows,5).Value)
  stresetup = Trim(objExcel.Cells(rows,1).Value)
  objTextFile.WriteLine(struid &" "& stresetup)
  
  IE.navigate "https://cn16.airwatchportals.com/AirWatch/#/User/List"
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Opened Airwatch website")
  t = IE.Document.title
  Set I = WScript.CreateObject("WScript.Shell")
  If (t="Login") Then
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Login into Airwatch")
  I.SendKeys "{ENTER}"
  WScript.Sleep 15000 
  End If
  WScript.Sleep 6000 
  c = 0
  Do While (c<28)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
  Loop
  I.SendKeys struid
  I.SendKeys "{ENTER}" 
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Search UID: "& struid)
  WScript.Sleep 4000
  I.SendKeys "{PRTSC}"

  
  rows = rows +1
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Moving to next eSetup")
Loop
objTextFile.WriteLine(Now())
objTextFile.WriteLine("Completed Departure in Airwatch Script")
objTextFile.Close
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
Msgbox "Completed. Check if accounts exist before closing each window !"