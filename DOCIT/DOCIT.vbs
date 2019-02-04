Option Explicit
Dim objExcel, objSheet, objSheet2, rows, col, row, words, a, x, objShell, objWorkbook, IE
Dim strname, stremail, strid, struid, strtype, I, c, strview
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
Set objWorkbook = objExcel.Workbooks.Open("C:\BMS Automation\DOCIT\data.xlsx")
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

rows = 1
words =0
col = 1
row = 1
x= 1

Do Until objExcel.Cells(rows,1).Value =  ""
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("eSetup :"& objExcel.Cells(rows,1).Value)

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
  strtype = Trim(objExcel.Cells(rows,11).Value)
  

  Set I = WScript.CreateObject("WScript.Shell")
  WScript.Sleep 1000
  IE.navigate "https://doccompliance.bms.com/" 
  WScript.Sleep 4000
  Do While (IE.Busy)
   WScript.Sleep 4000    
  Loop
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Opened DocCompliance website")
  WScript.Sleep 2000
  c = 0
  Do While (c<8)
   I.SendKeys "{TAB}"
   WScript.Sleep 200
   c = c+1
  Loop
  I.SendKeys "{DOWN}"
  WScript.Sleep 100
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys "{ENTER}"
  Do While (IE.Busy)    
     WScript.Sleep 100    
  Loop
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Login into DOCIT")
  WScript.Sleep 4000 
  c = 0
  Do While (c<3)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
  Loop
  I.SendKeys "{ENTER}"
  WScript.Sleep 5000
  c = 0
  Do While (c<19)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
  Loop
  I.SendKeys "{ENTER}"
  WScript.Sleep 3500 
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Entering User details")
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
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Selecting Company")
  I.SendKeys "{ENTER}"
  WScript.Sleep 1500
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys "{DOWN}"
  WScript.Sleep 100
  I.SendKeys "{DOWN}"
  WScript.Sleep 100
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys "{ENTER}"
  c = 0
  Do While (c<3)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
  Loop
  I.SendKeys "{ENTER}"
  WScript.Sleep 100
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys "{ENTER}"
  WScript.Sleep 3000
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  I.SendKeys Right("00000000" & strid, 8)
  WScript.Sleep 100
  I.SendKeys "{TAB}"
  WScript.Sleep 100
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Entering User Type")
  If (strtype = "QSD Viewer" OR strtype = "viewer" OR strtype = "view" OR strtype = "VIEWER") Then
   I.SendKeys "{DOWN}"
   WScript.Sleep 100
  End If
  c = 0
  Do While (c<4)
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   c = c+1
  Loop
  I.SendKeys "{ENTER}"
  WScript.Sleep 1500
  I.SendKeys "{ENTER}"
  WScript.Sleep 500
  I.SendKeys "{ENTER}"
  WScript.Sleep 7000
  I.SendKeys "{TAB}"

  If (strtype = "QSD Viewer" OR strtype = "View" OR strtype = "viewer" OR strtype = "VIEW") Then
   objTextFile.WriteLine(Now())
   objTextFile.WriteLine("User requesting for QSD Viewer Role")
   c = 0
   Do While (c<9)
    I.SendKeys "{Down}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   c=0
   Do While (c<4)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 100
   WScript.Sleep 2000
   c = 0
   Do While (c<5)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 7000
  ElseIf(strtype = "QSD Standard User" OR strtype = "STANDARD" OR strtype = "standard") Then
   objTextFile.WriteLine(Now())
   objTextFile.WriteLine("User requesting for QSD Standard User Role")
   c = 0
   Do While (c<7)
    I.SendKeys "{Down}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   c=0
   Do While (c<4)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{Down}"
   WScript.Sleep 100
   I.SendKeys "{Down}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   c=0
   Do While (c<4)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   c = 0
   Do While (c<3)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 7000
  ElseIf(strtype = "QSD Coordinator" OR strtype = "coordinator" OR strtype = "COORDINATOR") Then
   objTextFile.WriteLine(Now())
   objTextFile.WriteLine("User requesting for QSD Coordinator Role")
   c = 0
   Do While (c<7)
    I.SendKeys "{Down}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   
   c=0
   Do While (c<6)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   c = 0
   Do While (c<3)
    I.SendKeys "{Down}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 100
   c=0
   Do While (c<4)
    I.SendKeys "{TAB}"
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{Down}"
   WScript.Sleep 100
   I.SendKeys "{Down}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   c=0
   Do While (c<4)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   c = 0
   Do While (c<3)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 7000
  ElseIf(strtype = "QSD Super User" OR strtype = "super" OR strtype = "SUPER") Then
   objTextFile.WriteLine(Now())
   objTextFile.WriteLine("User requesting for QSD Super User Role")
   c = 0
   Do While (c<7)
    I.SendKeys "{Down}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 100
   c=0
   Do While (c<6)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   c = 0
   Do While (c<7)
    I.SendKeys "{Down}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 100
   c=0
   Do While (c<4)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{Down}"
   WScript.Sleep 100
   I.SendKeys "{Down}"
   WScript.Sleep 1000
   I.SendKeys "{TAB}"
   WScript.Sleep 100
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   c=0
   Do While (c<4)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 1000
   c = 0
   Do While (c<3)
    I.SendKeys "{TAB}"
    WScript.Sleep 100
    c = c+1
   Loop
   I.SendKeys "{ENTER}"
   WScript.Sleep 6000
  End If
  objTextFile.WriteLine(Now())
  objTextFile.WriteLine("Moving to check if there is another eSetup")
WScript.Sleep 6000

rows = rows + 1

Loop
objTextFile.WriteLine(Now())
objTextFile.WriteLine("Completed Script")
objTextFile.Close
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
Msgbox "End"