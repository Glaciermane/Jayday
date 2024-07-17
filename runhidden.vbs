Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell -ExecutionPolicy RemoteSigned -File run.ps1", 0
