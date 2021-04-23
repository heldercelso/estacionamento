Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c .\kivy_venv\Scripts\python.exe Parking.py"
oShell.Run strArgs, 0, false