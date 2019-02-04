Digital Certificate v2.0
=================================================
File list
- DigitalCertificate.vbs
- userdet.txt
- Groups.xlsx
=================================================
Functional Details:

- userid.txt
  Contains ids to add security groups
  Format = eSetup<space>UID<space>Role

- Groups.xlsx
  Contains list of roles and groups

- DigitalCertificate.vbs
  Adds security groups to user account on AD website. 
==================================================
Error handling
- No errors
- Can run multiple times with no negative effect on User accounts
- If mailgroup is required it automatically launches
==================================================
Limitiations
- Manual filling of userdet.txt
==================================================
Changelog

v2.0 - Removed hardcoding of roles and groups and transferred details to excel sheet. 