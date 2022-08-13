Option Explicit

Dim x
Set x = WScript.CreateObject("WScript.Network")

Dim strPath
strPath = "C:\Users\" & x.UserName & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"

Dim objShell
Set objShell = WScript.CreateObject("Shell.Application")
objShell.Explore strPath
Set objShell = Nothing
