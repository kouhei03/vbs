' ファイル概要: 終了コード分岐デモ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 152_ExitCodeBranchDemo.vbs
' 実行結果: 終了コード分岐デモ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, exitCode
Set shell=CreateObject("WScript.Shell")
' [処理解説] 外部プロセスを起動し、出力ストリームと終了コードを取得できるオブジェクトを受け取ります。
Set process=shell.Exec("cmd.exe /c exit 5")
Do While Not process.StdOut.AtEndOfStream
    process.StdOut.ReadLine
Loop
Do While Not process.StdErr.AtEndOfStream
    process.StdErr.ReadLine
Loop
exitCode=process.ExitCode
WScript.Echo "ExitCode=" & exitCode
If exitCode=0 Then
    WScript.Echo "Success branch."
Else
    WScript.Echo "Failure branch; this is an intentional demo exit code."
End If
