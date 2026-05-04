Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = shell.ExpandEnvironmentStrings("%LocalAppData%") & "\Microsoft\WindowsApps\pythonw.exe"
scriptPath = fso.BuildPath(scriptDir, "printer_gui.py")

shell.Run """" & pythonw & """ """ & scriptPath & """", 0, False
