SET p=C:\SphinxService\
SET s=Test4
REG ADD HKLM\SOFTWARE\INT /f /v WorkingDirectory /t REG_SZ /d %p%
sc create %s% binPath= %p%spnxsrv.exe
sc start %s%