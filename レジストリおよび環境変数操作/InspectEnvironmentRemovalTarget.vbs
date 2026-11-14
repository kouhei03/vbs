' ファイル概要: ユーザー環境変数の存在を確認する（削除なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectEnvironmentRemovalTarget.vbs <variable-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ユーザー環境変数の存在を確認する（削除なし）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, environment, variableName, entry, separator, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo InspectEnvironmentRemovalTarget.vbs <variable-name>"
    WScript.Quit 2
End If
variableName = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")
Set environment = shell.Environment("USER")
found = False
For Each entry In environment
    separator = InStr(entry, "=")
    If separator > 0 Then
        If StrComp(Left(entry, separator - 1), variableName, vbTextCompare) = 0 Then found = True
    End If
Next
If found Then
    WScript.Echo "Variable exists in the current user environment: " & variableName
Else
    WScript.Echo "Variable is not present in the current user environment: " & variableName
End If
WScript.Echo "No environment variable was removed."
