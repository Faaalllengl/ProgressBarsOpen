Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
server = folder & "\server.py"

shell.CurrentDirectory = folder

If TestCommand(shell, "py") Then
  shell.Run "py """ & server & """", 1, True
ElseIf TestCommand(shell, "python") Then
  shell.Run "python """ & server & """", 1, True
Else
  MsgBox "Python ne ustanovlen." & vbCrLf & vbCrLf & "Skachayte: https://www.python.org/downloads/" & vbCrLf & "Vklyuchite: Add python.exe to PATH", vbCritical, "Progress Bar"
End If

Function TestCommand(sh, cmd)
  On Error Resume Next
  Dim code
  code = sh.Run("cmd /c " & cmd & " --version >nul 2>&1", 0, True)
  TestCommand = (code = 0)
  On Error GoTo 0
End Function
