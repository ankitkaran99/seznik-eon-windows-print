Set oShell = CreateObject("WScript.Shell")
pythonExe = "C:\Users\ankit\OneDrive\Desktop\seznik-eon-printer-toolkit\python\tools\python.exe"
relayScript = "C:\Users\ankit\OneDrive\Desktop\seznik-eon-printer-toolkit\printer_relay.py"
cmd = Chr(34) & pythonExe & Chr(34) & " " & Chr(34) & relayScript & Chr(34) & " --host 127.0.0.1 --port 9100"
oShell.CurrentDirectory = "C:\Users\ankit\OneDrive\Desktop\seznik-eon-printer-toolkit"
oShell.Run cmd, 0, False
