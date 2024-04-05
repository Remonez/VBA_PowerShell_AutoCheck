Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -ExecutionPolicy Bypass -File D:\Downloads\RunExcelMacro.ps1", 0, False
Set objShell = Nothing

