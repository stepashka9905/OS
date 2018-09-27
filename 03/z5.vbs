dim WshShell
set WshShell = WScript.CreateObject("Wscript.Shell")
WshShell.Run "notepad.exe " & WScript.ScriptFullName,1,true