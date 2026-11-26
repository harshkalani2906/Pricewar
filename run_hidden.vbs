' Starts RUN IN BACKGROUND.bat with no black window (used by the automatic runs).
Set sh = CreateObject("WScript.Shell")
dir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dir
a = ""
For Each x In WScript.Arguments
  a = a & " " & x
Next
sh.Run """" & dir & "\RUN IN BACKGROUND.bat""" & a, 0, True
