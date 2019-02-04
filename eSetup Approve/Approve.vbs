Option Explicit
Dim objExcel, objExcel1, objShell, objWorkbook, objPDHQ, strExcelPath
Dim strPDHQ, objSheet, objSheet2, rows, col, row, words, a, x, b, text
Dim objbook, IE, intRow, comments, WshShell, element, htmlelement

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

Set objExcel = CreateObject("Excel.Application")
Set objbook = objExcel.Workbooks.Open("C:\BMS Automation\eSetup Approve\data.xlsx")
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

Set objExcel1 = CreateObject("Excel.Application")
Set objbook = objExcel1.Workbooks.Open("C:\BMS Automation\eSetup Approve\comments.xlsx")
Set objSheet2 = objExcel1.ActiveWorkbook.Worksheets(1)

rows = 1
words =0
col = 1


Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = 1

Do Until objExcel.Cells(rows,1).Value =  ""
 row = 1
 WScript.Sleep 1000
 IE.navigate objExcel.Cells(rows,2).Value
 Do While (IE.Busy)    
    WScript.Sleep 100    
 Loop
 Do Until objExcel1.Cells(row,1).Value =  ""
  If objExcel1.Cells(row,1).Value = objExcel.Cells(rows,1).Value Then
   Err.Clear
   On Error Resume Next
   IE.Document.All.Item("variables.user.provNotes").Value = objExcel1.Cells(row,2).Value
   If Err.Number <> 0 Then
    IE.Document.All.Item("variables.user.myComments").Value = objExcel1.Cells(row,2).Value
   End If
   objTextFile.WriteLine(Now())
   objTextFile.WriteLine("Added Comment: "& objExcel1.Cells(row,2).Value)
   On Error GoTo 0
  End If
  row = row +1
  WScript.Sleep 40   
 Loop
 IE.Document.getElementById("Approve").Click
 objTextFile.WriteLine(Now())
 objTextFile.WriteLine("Approved eSetup")
 Do While (IE.Busy)    
    WScript.Sleep 100    
 Loop
 rows = rows+1
 WScript.Sleep 1000
Loop

objExcel1.ActiveWorkbook.Close
objExcel1.Application.Quit
objExcel1.Quit
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
Msgbox "Completed Approving"
