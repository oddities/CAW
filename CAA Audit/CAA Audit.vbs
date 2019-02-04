Option Explicit
Dim objExcel, objExcel1, objShell, objWorkbook, objPDHQ, strExcelPath
Dim objSheet, objSheet2, rows, col, row, words, a, x, b, c, I
Dim objbook, IE, intRow, stresetup

Set objShell = CreateObject("Wscript.Shell")
objShell.Run("taskkill /im iexplore.exe"), 1, TRUE

Set objExcel = CreateObject("Excel.Application")
Set objbook = objExcel.Workbooks.Open("C:\BMS Automation\CAA Audit\data.xlsx")
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

x=1
rows = 1
words =0
col = 1

Do Until objExcel.Cells(rows,1).Value =  ""

 Set IE = CreateObject("InternetExplorer.Application")
 IE.Visible = 1
 stresetup = objExcel.Cells(rows,1).Value
 IE.navigate "http://esetup.bms.com/esetup2-web/home.jsf"
 Do While (IE.Busy)
   WScript.Sleep 4000    
 Loop
 Set I = WScript.CreateObject("WScript.Shell")
 c = 0
 Do While (c<13)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
 Loop
 I.SendKeys "{ENTER}"
 WScript.Sleep 1000
 c = 0
 Do While (c<9)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
 Loop
 I.SendKeys stresetup
 WScript.Sleep 100
 I.SendKeys "{TAB}"
 WScript.Sleep 10
 I.SendKeys "{ENTER}"
 WScript.Sleep 1000
 c = 0
 Do While (c<10)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
 Loop
 Do While (IE.Busy)
   WScript.Sleep 4000    
 Loop
 I.SendKeys "{ENTER}"
 
rows = rows + 1
Loop
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
Msgbox "completed"