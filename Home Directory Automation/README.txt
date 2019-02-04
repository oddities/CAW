HOME FOLDER AUTOMATION v2.0
=================================================
File list
- ADScript.ps1
- ADAccess.vbs
- data.xls
- userid.txt
=================================================
Functional Details:

- userid.txt
  Contains list of BMS UID

- data.xls
  Altomatically generated

- ADScript.ps1
  Adds SCRIPT.VBS to listed UID in userid.txt

- ADAccess.vbs
  Adds security groups to user account on AD website.
==================================================
Error rate
- No errors
- Can run multiple times with no negative effect on User accounts
==================================================
Limitiations
- Manual filling of userid.txt
- 2 separate scripts
==================================================
Changelog
-Eliminated manual filling of details