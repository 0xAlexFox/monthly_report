Option Explicit

Dim fso, shell, baseDir, pyw, scriptPath, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
pyw = baseDir & "\.venv\Scripts\pythonw.exe"
scriptPath = baseDir & "\compare_payments.py"

If Not fso.FileExists(pyw) Then
    MsgBox ".venv is not found." & vbCrLf & _
           "Run setup first:" & vbCrLf & _
           "powershell -ExecutionPolicy Bypass -File .\setup.ps1", _
           vbCritical, "Bank Reconciliation"
    WScript.Quit 1
End If

cmd = """" & pyw & """ """ & scriptPath & """ --gui"
shell.Run cmd, 0, False
