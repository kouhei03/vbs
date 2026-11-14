' 概要: 子プロセスに継承されないプロセススコープ環境変数と作業ディレクトリの例です。
' 実行方法: Usage: cscript //nologo 089_ProcessEnvironmentDemo.vbs <existing-folder>
' 処理結果: プロセス環境変数と作業ディレクトリの例。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, shell, env, folder
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 089_ProcessEnvironmentDemo.vbs <existing-folder>"
    WScript.Quit 2
End If
' 主処理: プロセス環境変数と作業ディレクトリの例。
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FolderExists(folder) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("PROCESS")
env("VBS_DEMO_VALUE") = "temporary-process-value"
shell.CurrentDirectory = folder
WScript.Echo "VBS_DEMO_VALUE=" & env("VBS_DEMO_VALUE")
WScript.Echo "CurrentDirectory=" & shell.CurrentDirectory
