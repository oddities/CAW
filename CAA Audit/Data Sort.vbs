Option Explicit
Dim objExcel, objExcel1, objShell, objWorkbook, objPDHQ, strExcelPath
Dim objSheet, objSheet2, rows, col, row, words, a, x, b, c, I
Dim objbook, IE, intRow, stresetup
Dim max,min
max=7000
min=1

Set objShell = CreateObject("Wscript.Shell")
objShell.Run("taskkill /im iexplore.exe"), 1, TRUE

Set objExcel = CreateObject("Excel.Application")
Set objbook = objExcel.Workbooks.Open("C:\BMS Automation\CAA Audit\data.xlsx")
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

Set objExcel1 = CreateObject("Excel.Application")
Set objbook = objExcel1.Workbooks.Open("C:\BMS Automation\CAA Audit\CAA_Esetup_Complete.xls")

x=1
rows = 1
words =0
col = 1

Do Until x = 31
 
 Randomize
 rows= Int((max-min+1)*Rnd+min)
 If (objExcel1.Cells(rows,6).Value ="Manoj Kumar Mallick" OR objExcel1.Cells(rows,6).Value ="Pushpa Pujitha V" OR objExcel1.Cells(rows,6).Value ="Kalpana Andhavarapu" OR objExcel1.Cells(rows,6).Value ="Manu Saraswat" OR objExcel1.Cells(rows,6).Value ="Dipendu Patar" OR objExcel1.Cells(rows,6).Value ="Aditya Kumar Khankhoje") Then
  objExcel.Cells(x,1).Value = objExcel1.Cells(rows,3).Value
  objExcel.Cells(x,2).Value = objExcel1.Cells(rows,19).Value
  objExcel.Cells(x,3).Value = objExcel1.Cells(rows,6).Value
  objExcel.Cells(x,4).Value = objExcel1.Cells(rows,1).Value
  objExcel.Cells(x,5).Value = objExcel1.Cells(rows,10).Value
  objExcel.ActiveWorkbook.Save
  x = x+1
 End If
 
Loop
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
objExcel1.ActiveWorkbook.Close
objExcel1.Application.Quit
objExcel1.Quit
Msgbox "Completed"