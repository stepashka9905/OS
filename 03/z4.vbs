Dim objFSO, out
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set out = objFSO.CreateTextFile("z4.bat", True)
out.Write "start """" ""C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OpenOffice 4.1.2\OpenOffice Calc"""
out.Close
