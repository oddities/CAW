Option Explicit
Dim objExcel, objPDHQ, strExcelPath, strPDHQ, objSheet, objSheet2, rows, col, row, words, a, c, x, IE, objShell, objWorkbook
Dim struid, strspec, strsite, stresetup, I, t

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
  
  struid = Trim(objExcel.Cells(rows,5).Value)
  stresetup = Trim(objExcel.Cells(rows,1).Value)
  
  
  IE.navigate "https://cn16.airwatchportals.com/AirWatch/#/User/List"
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  t = IE.Document.title
  Set I = WScript.CreateObject("WScript.Shell")
  If (t="Login") Then
  
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
  WScript.Sleep 4000
  I.SendKeys "{PRTSC}"

  
  rows = rows +1
Loop

objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
Msgbox "Completed. Check if accounts exist before closing each window !"