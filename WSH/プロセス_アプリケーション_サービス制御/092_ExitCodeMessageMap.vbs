' 概要: 終了コードを定義済みメッセージへ対応づけます（外部コマンドは実行しません）。
' 実行方法: Usage: cscript //nologo 092_ExitCodeMessageMap.vbs <exit-code>
' 処理結果: 終了コードから定義済みメッセージへの対応づけ。処理結果や状態はコンソールに表示されます。
' コード解説: スクリプト引数を使って対象を特定し、検証後に処理します。
Option Explicit
Dim code
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    ' 取得した値または処理状態をコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 092_ExitCodeMessageMap.vbs <exit-code>"
    WScript.Quit 2
End If
' 主処理: 終了コードから定義済みメッセージへの対応づけ。
If Not IsNumeric(WScript.Arguments(0)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < -32768 Or CDbl(WScript.Arguments(0)) > 32767 Then WScript.Quit 2
code = CLng(WScript.Arguments(0))
Select Case code
    Case 0: WScript.Echo "Success: operation completed."
    Case 1: WScript.Echo "General failure: inspect the operation log."
    Case 2: WScript.Echo "Invalid arguments or configuration."
    Case 3: WScript.Echo "Requested item was not found."
    Case Else: WScript.Echo "Unmapped exit code: " & code
End Select
