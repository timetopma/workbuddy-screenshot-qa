Option Explicit
' Launch the floating screenshot button without any console window.
' No hardcoded paths: pythonw is looked up on PATH first, then in the usual
' per-user / machine-wide Python install folders, so the folder stays portable.

Dim fso, sh, appdir, py, script, dirs, d, p, i, vers
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

appdir = fso.GetParentFolderName(WScript.ScriptFullName)
script = appdir & "\screenshot_qa.py"

py = ""

' 1) PATH: pythonw.exe, or the py launcher's windowed build (pyw.exe)
dirs = Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
For i = 0 To UBound(dirs)
    d = Trim(dirs(i))
    If py = "" And Len(d) > 0 Then
        If fso.FileExists(d & "\pythonw.exe") Then py = d & "\pythonw.exe"
        If py = "" And fso.FileExists(d & "\pyw.exe") Then py = d & "\pyw.exe"
    End If
Next

' 2) per-user installs, newest version first
If py = "" Then
    vers = Array("314", "313", "312", "311", "310", "39")
    For i = 0 To UBound(vers)
        If py = "" Then
            p = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & _
                "\Programs\Python\Python" & vers(i) & "\pythonw.exe"
            If fso.FileExists(p) Then py = p
        End If
    Next
End If

' 3) machine-wide installs
If py = "" Then
    vers = Array("313", "312", "311", "310", "39")
    For i = 0 To UBound(vers)
        If py = "" Then
            p = "C:\Program Files\Python" & vers(i) & "\pythonw.exe"
            If fso.FileExists(p) Then py = p
        End If
    Next
End If

If py = "" Then
    MsgBox "Python 3 was not found." & vbCrLf & vbCrLf & _
           "Install it from https://www.python.org/downloads/ and tick " & _
           """Add python.exe to PATH"" during setup.", vbExclamation, "WorkBuddy Screenshot"
    WScript.Quit 1
End If

sh.CurrentDirectory = appdir
sh.Run """" & py & """ """ & script & """", 0, False
