Set objShell = CreateObject("WScript.Shell")
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
objShell.Run "powershell.exe -NoProfile -File """ & strPath & "\always_on_top.ps1""", 0, False
