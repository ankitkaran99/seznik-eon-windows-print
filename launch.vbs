Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(scriptDir, "printer_gui.py")
pythonw = shell.ExpandEnvironmentStrings("%SEZNIK_EON_PYTHONW%")

If pythonw = "%SEZNIK_EON_PYTHONW%" Or pythonw = "" Then
    cmd = "cmd /c pyw -3 " & Chr(34) & scriptPath & Chr(34)
Else
    cmd = Chr(34) & pythonw & Chr(34) & " " & Chr(34) & scriptPath & Chr(34)
End If

shell.Run cmd, 0, False
