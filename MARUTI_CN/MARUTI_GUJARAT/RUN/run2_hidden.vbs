Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

runPath = fso.GetParentFolderName(WScript.ScriptFullName)
rootPath = fso.GetParentFolderName(runPath)
systemPath = rootPath & "\SYSTEM"

scriptPath = systemPath & "\watchM_input.py"

' Try venv first
pythonExe = systemPath & "\venv\Scripts\python.exe"

If Not fso.FileExists(pythonExe) Then
    pythonExe = "python"
End If

cmd = "cmd /c cd /d """ & systemPath & """ && """ & pythonExe & """ """ & scriptPath & """"

WshShell.Run cmd, 1, False