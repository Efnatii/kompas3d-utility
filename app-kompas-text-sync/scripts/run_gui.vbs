Option Explicit

Dim fso
Dim shell
Dim appRoot
Dim launcherExe
Dim guiScript
Dim command

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appRoot = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
launcherExe = fso.BuildPath(appRoot, "bin\app-kompas-text-sync.exe")
guiScript = fso.BuildPath(appRoot, "scripts\gui_sync.ps1")

If fso.FileExists(launcherExe) Then
    command = """" & launcherExe & """"
    shell.Run command, 0, False
    WScript.Quit 0
End If

If Not fso.FileExists(guiScript) Then
    MsgBox "gui_sync.ps1 not found: " & guiScript, vbCritical, "kompas-text-sync app"
    WScript.Quit 2
End If

command = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & guiScript & """"
shell.Run "powershell.exe " & command, 0, False
