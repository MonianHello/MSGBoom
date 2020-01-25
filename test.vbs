Set Wshshell=CreateObject("Wscript.Shell") 
Wscript.sleep 3000
dim s 
do until s=2000
s=s+1 
WshShell.SendKeys "^v^{ENTER}"
loop
