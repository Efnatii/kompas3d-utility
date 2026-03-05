Option Explicit

Dim fso
Dim shell
Dim appRoot
Dim launcherExe
Dim runGuiVbs
Dim command

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appRoot = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
launcherExe = fso.BuildPath(appRoot, "bin\app-kompas-text-sync.exe")
runGuiVbs = fso.BuildPath(appRoot, "scripts\run_gui.vbs")

If fso.FileExists(launcherExe) Then
    command = """" & launcherExe & """"
    shell.Run command, 0, False
    WScript.Quit 0
End If

If Not fso.FileExists(runGuiVbs) Then
    MsgBox "run_gui.vbs not found: " & runGuiVbs, vbCritical, "kompas-text-sync app"
    WScript.Quit 2
End If

command = "wscript.exe """ & runGuiVbs & """"
shell.Run command, 0, False
