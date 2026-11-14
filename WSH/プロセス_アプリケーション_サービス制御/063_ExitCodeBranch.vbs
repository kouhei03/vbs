' 概要: 固定コマンドの終了コードを取得し、成功／失敗の分岐例を表示します。
' 実行方法: cscript //nologo 063_ExitCodeBranch.vbs
' 処理結果: 固定コマンドの終了コード分岐。処理結果や状態はコンソールに表示されます。
' コード解説: WScript.Shellでコマンドを実行し、終了コードに応じて結果を処理します。
Option Explicit
Dim shell, exitCode
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドを実行し、終了コードを取得します。
exitCode = shell.Run("cmd.exe /c ver", 1, True)
If exitCode = 0 Then
    WScript.Echo "Success: Windows version query completed."
Else
    WScript.Echo "Failure: fixed command returned " & exitCode
    WScript.Quit exitCode
End If
