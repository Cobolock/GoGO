SET s=Test4
sc stop %s%
sleep 1
taskkill /IM spnxsrv.exe
REG DELETE HKLM\SOFTWARE\INT /f
sc delete %s%