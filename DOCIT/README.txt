DOCIT v1.0
=================================================
File list
- pdhq.vbs
- userid.txt
=================================================
Functional Details:

- userid.txt
  Contains ids to add security groups
  Format = First name<space>Last name<space>Email<space>BMSID<space>Role

- DOCIT.vbs
  Creates new user account in DOCIT.  
==================================================
Error handling
- Skips to next user if there are any errors.
- First launch may cause failure in first user although its rare and depends on current system status
- Can run multiple times with no negative effect on User accounts
- Kills old existing internet explorer windows before start as a fresh start
==================================================
Limitiations
- Manual filling of userdet.txt
==================================================
