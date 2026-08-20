Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = appDir & "\venv\Scripts\pythonw.exe"
launcher = appDir & "\run_app.pyw"

If Not fso.FileExists(pythonw) Then
    pythonw = "pythonw.exe"
End If

shell.Run """" & pythonw & """ """ & launcher & """", 0, False
