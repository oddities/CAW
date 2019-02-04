Option Explicit
Dim objExcel, objSheet, objSheet2, rows, col, row, words, a, x, objShell, objWorkbook, IE
Dim strname, stremail, strid, struid, strtype, I, c, strview, strcab, strrole

Set objShell = CreateObject("Wscript.Shell")
objShell.Run("taskkill /im iexplore.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
WScript.Sleep 2000

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\BMS Automation\DOCMAN\data.xlsx")
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

rows = 1
words =0
col = 1
row = 1
x= 1

Do Until objExcel.Cells(rows,1).Value =  ""
 
  
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = 1
  strname = Trim(objExcel.Cells(rows,2).Value)

  If (objExcel.Cells(rows,3).Value = "") Then
   stremail = "user@bms.com"
  Else
   stremail = Trim(objExcel.Cells(rows,3).Value)
  End If

  strid = Right(objExcel.Cells(rows,4).Value,8)
  struid = Trim(objExcel.Cells(rows,5).Value)
  strtype = Trim(objExcel.Cells(rows,9).Value)
  strrole = Trim(objExcel.Cells(rows,12).Value)
  strcab = Trim(objExcel.Cells(rows,11).Value)

  Set I = WScript.CreateObject("WScript.Shell")
  WScript.Sleep 1000
  IE.navigate "https://doccompliance.bms.com/" 
  WScript.Sleep 3000
  Do While (IE.Busy)
   WScript.Sleep 4000    
  Loop
  WScript.Sleep 2000
  c = 0
  Do While (c<8)
   I.SendKeys "{TAB}"
   WScript.Sleep 200
   c = c+1
  Loop
  I.SendKeys "{UP}"
  WScript.Sleep 100
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys "{ENTER}"
  'IE.Document.getElementsByName("btnLogin").Item(0).Click    
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  WScript.Sleep 4000
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys "{TAB}"
  WScript.Sleep 100 
  I.SendKeys "{ENTER}" 'Click Admin Tab here
  WScript.Sleep 5000 
  c = 0
  Do While (c<19)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
  Loop 
  I.SendKeys "{ENTER}" 'Click New User
  WScript.Sleep 3500 
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys struid
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys strname
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys stremail
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys Right("00000000" & strid, 8)
  c = 0
  Do While (c<3)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
  Loop
  I.SendKeys "{ENTER}" 
  WScript.Sleep 3500
  If (strcab = "Accenture (ACN" OR strcab = "ACN") Then
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{DOWN}"
   WScript.Sleep 100
   I.SendKeys "{DOWN}"
   WScript.Sleep 100
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 100
  End If
  c = 0
  Do While (c<3)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
  Loop
  I.SendKeys "{ENTER}"
  I.SendKeys "{TAB}"
  I.SendKeys "{ENTER}"
  WScript.Sleep 3000
  I.SendKeys "{TAB}"
  I.SendKeys Right("00000000" & strid, 8)
  c = 0
  Do While (c<5)
   I.SendKeys "{TAB}"
   c = c+1
  Loop
  I.SendKeys "{ENTER}"
  WScript.Sleep 1500
  I.SendKeys "{ENTER}"
  WScript.Sleep 500
  I.SendKeys "{ENTER}"
  WScript.Sleep 1000
  I.SendKeys "{TAB}"
  If (strrole = "Viewer" OR strrole = "viewer" OR strrole = "View" OR strrole ="view") Then
   c = 0
   Do While (c<5)
    I.SendKeys "{Down}"
	WScript.Sleep 100
    c = c+1
   Loop 
   I.SendKeys "{TAB}"
   WScript.Sleep 1000
   I.SendKeys "{ENTER}"
   c = 0
   Do While (c<23)
    I.SendKeys "{Down}"
	WScript.Sleep 100
    c = c+1
   Loop 
   I.SendKeys "{TAB}"
   WScript.Sleep 1000
   I.SendKeys "{ENTER}"
   c = 0
   Do While (c<4)
    I.SendKeys "{TAB}"
	WScript.Sleep 100
    c = c+1
   Loop 
   I.SendKeys "{ENTER}"
   WScript.Sleep 3000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
  End If
rows = rows + 1
WScript.Sleep 3000
Loop

objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
WScript.Sleep 2000
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
WScript.Sleep 2000
Msgbox "completed"