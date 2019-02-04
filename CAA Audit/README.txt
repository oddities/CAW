=============================================================
CAA Audit Automation
=============================================================
Path in VDI where the folder has to be placed

C:\BMS Automation\
=============================================================
Files:
- CAA Audit.exe
- Data Sort 30.exe
- Data Sort 50.exe
- data.xlsx
- CAA_Esetup_Complete.xlsx
=============================================================
Functional Details:

- CAA_Esetup_Complete.xlsx
  Contains the data filtered from the eSetup report meant for audit.

- Data Sort 30/50.exe
  Takes random 30 or 50 esetup ticket details from CAA_Esetup_Complete.xlsx sheet and fills the data.xlsx excel file

- data.xlsx
  Contains detailed information of eSetup ticket like Ticket number, Analyst name, eSetup name and date completed

- CAA Audit.exe
  Opens the eSetup ticket individually from the data.xlsx sheet
==============================================================
Procedure
- Paste the "CAA_Esetup_Complete.xlsx" with the data filtered from the eSetup report.
- If you are auditing 30 tickets run "Data Sort 30.exe" or for 50 tickets run "Data Sort 50.exe"
- After the Sort application completes running you will get Random 30 / 50 tickets in the "data.xlsx" file.
- You can use this data to fill the Audit sheet which has to be uploaded to the Audit sharepoint.
- Now Run "CAA Audit.exe" application to open each eSetup on a different Internet Explorer window to individually audit them.
