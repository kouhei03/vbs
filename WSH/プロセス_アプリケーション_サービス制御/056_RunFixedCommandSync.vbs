' 概要: 固定のWindowsバージョン照会コマンドを同期実行し、終了コードを表示します。
' 実行方法: cscript //nologo 056_RunFixedCommandSync.vbs
' 処理結果: 固定コマンドの同期実行。処理結果や状態はコンソールに表示されます。
' コード解説: WScript.Shellでコマンドを実行し、終了コードに応じて結果を処理します。
Option Explicit
Dim shell, exitCode
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドを実行し、終了コードを取得します。
exitCode = shell.Run("cmd.exe /c ver", 1, True)
WScript.Echo "ExitCode=" & exitCode
If exitCode <> 0 Then WScript.Quit exitCode
