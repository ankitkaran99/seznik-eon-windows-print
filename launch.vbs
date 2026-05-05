Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(scriptDir, "printer_gui.py")
shell.CurrentDirectory = scriptDir

Function FileExists(path)
    FileExists = False
    If Len(path) = 0 Then
        Exit Function
    End If

    On Error Resume Next
    FileExists = fso.FileExists(path)
    On Error GoTo 0
End Function

Function CanRunGui(pythonExe)
    Dim checkCmd

    CanRunGui = False
    If Len(pythonExe) = 0 Then
        Exit Function
    End If

    checkCmd = Chr(34) & pythonExe & Chr(34) & " -c " & Chr(34) & "import tkinter" & Chr(34)
    On Error Resume Next
    CanRunGui = (shell.Run(checkCmd, 0, True) = 0)
    On Error GoTo 0
End Function

Function ChoosePythonConsole()
    Dim candidates, candidate
    candidates = Array( _
        shell.ExpandEnvironmentStrings("%SEZNIK_EON_PYTHON%"), _
        "python" _
    )

    For Each candidate In candidates
        If FileExists(candidate) Or candidate = "python" Then
            If CanRunGui(candidate) Then
                ChoosePythonConsole = candidate
                Exit Function
            End If
        End If
    Next

    ChoosePythonConsole = ""
End Function

Function ChoosePythonWindowed(pythonConsole)
    Dim parentDir, pythonwPath

    If Len(pythonConsole) = 0 Then
        ChoosePythonWindowed = ""
        Exit Function
    End If

    If LCase(pythonConsole) = "python" Then
        ChoosePythonWindowed = "pythonw"
        Exit Function
    End If

    parentDir = fso.GetParentFolderName(pythonConsole)
    pythonwPath = fso.BuildPath(parentDir, "pythonw.exe")
    If FileExists(pythonwPath) Then
        ChoosePythonWindowed = pythonwPath
    Else
        ChoosePythonWindowed = pythonConsole
    End If
End Function

pythonConsole = ChoosePythonConsole()
If Len(pythonConsole) = 0 Then
    MsgBox "Unable to launch the GUI." & vbCrLf & vbCrLf & _
        "No Python installation with tkinter support was found." & vbCrLf & _
        "The selected system Python does not appear to include tkinter." & vbCrLf & vbCrLf & _
        "Install a full Python build with tkinter, or run configure_relay_printer.ps1 again so it can repair the launcher environment.", _
        vbCritical, "Seznik EON Printer Toolkit"
    WScript.Quit 1
End If

pythonWindowed = ChoosePythonWindowed(pythonConsole)
cmd = Chr(34) & pythonWindowed & Chr(34) & " " & Chr(34) & scriptPath & Chr(34)
shell.Run cmd, 0, False
