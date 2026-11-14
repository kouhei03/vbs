' 子プロセスに継承されないプロセススコープ環境変数と作業ディレクトリの例です。
Option Explicit
Dim fso, shell, env, folder
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 089_ProcessEnvironmentDemo.vbs <existing-folder>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FolderExists(folder) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("PROCESS")
env("VBS_DEMO_VALUE") = "temporary-process-value"
shell.CurrentDirectory = folder
WScript.Echo "VBS_DEMO_VALUE=" & env("VBS_DEMO_VALUE")
WScript.Echo "CurrentDirectory=" & shell.CurrentDirectory
