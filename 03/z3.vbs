Dim path
Set objFSO = CreateObject("Scripting.FileSystemObject")
path = InputBox("Vvedite path to file: ")
objFSO.CopyFile path, Replace(path,".","-copy."), true
