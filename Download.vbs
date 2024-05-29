Option Explicit

Dim objShell, strCommand

Set objShell = CreateObject("WScript.Shell")

strCommand = "powershell.exe -Command ""Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/3U4NL0UBS3R/VBS_Script_Download_Test/main/Error.vbs' -OutFile 'C:\Users\euanl\Downloads\Download.vbs'"""

objShell.Run strCommand, 0, True

Set objShell = Nothing
