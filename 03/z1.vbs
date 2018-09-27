Dim WShShell, path 
set WshShell = WScript.CreateObject("WScript.Shell") 
path = InputBox("Vvedite path to file: ")
WshShell.Run "notepad" &path
WshShell.AppActivate "notepad"
