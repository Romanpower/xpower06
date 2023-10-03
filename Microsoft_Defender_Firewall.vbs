' Open Notepad
Set objShell = CreateObject("WScript.Shell")
objShell.Run "notepad.exe"

' Wait for Notepad to open
Do While Not objShell.AppActivate("ABDUL - Notepad")
    Wscript.Sleep 100


' Send the text to Notepad
objShell.SendKeys "YOU ARE HACKER BY ROMAN"
Loop
