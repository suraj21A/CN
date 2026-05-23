Dim WshShell
Dim fso
Dim rootDir
Dim systemDir
Dim pythonPath
Dim command

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

rootDir = fso.GetParentFolderName(WScript.ScriptFullName)
rootDir = fso.GetParentFolderName(rootDir)

systemDir = rootDir & "\SYSTEM"
pythonPath = systemDir & "\venv\Scripts\python.exe"

If Not fso.FileExists(pythonPath) Then
    pythonPath = "python"
End If

command = "cmd.exe /k cd /d " & Chr(34) & systemDir & Chr(34) & _
          " && " & Chr(34) & pythonPath & Chr(34) & " watchR_input.py"

WshShell.Run command, 1, False