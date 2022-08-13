Option Explicit
On Error Resume Next

Dim objShell

Set objShell = WScript.CreateObject("Shell.Application")
If Err.Number = 0 Then
	    objShell.FileRun
	Else
		WScript.Echo "ÉGÉâÅ[: " & Err.Description
		End If

		Set objShell = Nothing'

