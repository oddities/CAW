Option Explicit
Dim objExcel, objPDHQ, strExcelPath, strPDHQ, objSheet, objSheet2, rows, col, row, words, a, x, IE, objShell, objWorkbook
Dim struid, strspec, strsite, stresetup
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
Set objWorkbook = objExcel.Workbooks.Open("C:\BMS Automation\Devlims\data.xlsx")
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
  strspec = Trim(objExcel.Cells(rows,15).Value)
  strsite  = Trim(objExcel.Cells(rows,13).Value)
  stresetup = Trim(objExcel.Cells(rows,1).Value)
  objTextFile.WriteLine(struid &" "& stresetup &" "& strsite)
  
  IE.navigate "https://adaccounts.ads.bms.com/showUser.aspx?uid="+struid
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Opened AD Website")
  IE.Document.All.Item("tb_eSetupControl").Value =stresetup
  If (strspec = "Production" OR strspec = "production") Then
   If (strsite = "HPW" OR strsite = "BLY") Then
    
    IE.Document.All.Item("tb_manualGroupName").Value = "GG_IMSS_BPPD_DEVLIMS_PROD_APP_HPW_USERS"
    IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
	objTextFile.WriteLine(Now())
    objTextFile.WriteLine("Added Group : GG_IMSS_BPPD_DEVLIMS_PROD_APP_HPW_USERS")

   ElseIf (strsite = "DVN" OR strsite = "HOP" OR strsite = "SLE" OR strsite = "SYR") Then

    IE.Document.All.Item("tb_manualGroupName").Value = "GG_IMSS_BPPD_DEVLIMS_PROD_APP_SYR_USERS"
    IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
	objTextFile.WriteLine(Now())
    objTextFile.WriteLine("Added Group : GG_IMSS_BPPD_DEVLIMS_PROD_APP_SYR_USERS")
   End If

  ElseIf (strspec = "Test" OR strspec = "test") Then
   If (strsite = "HPW" OR strsite = "BLY") Then
 
    IE.Document.All.Item("tb_manualGroupName").Value = "GG_IMSS_BPPD_DEVLIMS_TEST_APP_HPW_USERS"
    IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
    objTextFile.WriteLine(Now())
    objTextFile.WriteLine("Added Group : GG_IMSS_BPPD_DEVLIMS_TEST_APP_HPW_USERS")
   ElseIf (strsite = "DVN" OR strsite = "HOP" OR strsite = "SLE" OR strsite = "SYR") Then

    IE.Document.All.Item("tb_manualGroupName").Value = "GG_IMSS_BPPD_DEVLIMS_TEST_APP_SYR_USERS"
    IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
	objTextFile.WriteLine(Now())
    objTextFile.WriteLine("Added Group : GG_IMSS_BPPD_DEVLIMS_TEST_APP_SYR_USERS")
   End If

  ElseIf (strspec = "Development" OR strspec = "development") Then
   If (strsite = "HPW" OR strsite = "BLY") Then
 
    IE.Document.All.Item("tb_manualGroupName").Value = "GG_IMSS_BPPD_DEVLIMS_DEV_APP_HPW_USERS"
    IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
	objTextFile.WriteLine(Now())
    objTextFile.WriteLine("Added Group : GG_IMSS_BPPD_DEVLIMS_DEV_APP_HPW_USERS")

   ElseIf (strsite = "DVN" OR strsite = "HOP" OR strsite = "SLE" OR strsite = "SYR") Then

    IE.Document.All.Item("tb_manualGroupName").Value = "GG_IMSS_BPPD_DEVLIMS_DEV_APP_SYR_USERS"
    IE.Document.getElementsByName("b_manualGroupName").Item(0).Click
	objTextFile.WriteLine(Now())
    objTextFile.WriteLine("Added Group : GG_IMSS_BPPD_DEVLIMS_DEV_APP_SYR_USERS")
   End If

  End If
  rows = rows+1
Loop

objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit