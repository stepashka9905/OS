copy "" z1.vbs
echo Dim WShShell > z1.vbs
echo set WshShell = WScript.CreateObject("WScript.Shell") >> z1.vbs
echo WshShell.Run "notepad" >> z1.vbs
