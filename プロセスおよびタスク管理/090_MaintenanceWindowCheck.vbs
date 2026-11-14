' ファイル概要: 保守時間帯の判定のみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 090_MaintenanceWindowCheck.vbs <start-minute-0-1439> <end-minute-0-1439>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 保守時間帯の判定のみ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim startMinute, endMinute, nowMinute, insideWindow
If WScript.Arguments.Count <> 2 Then WScript.Echo "Usage: cscript //nologo 090_MaintenanceWindowCheck.vbs <start-minute-0-1439> <end-minute-0-1439>": WScript.Quit 2
If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Echo "Times must be minute numbers.": WScript.Quit 2
startMinute = CLng(WScript.Arguments(0))
endMinute = CLng(WScript.Arguments(1))
If startMinute < 0 Or startMinute > 1439 Or endMinute < 0 Or endMinute > 1439 Then WScript.Echo "Minutes must be between 0 and 1439.": WScript.Quit 2
nowMinute = Hour(Now) * 60 + Minute(Now)
If startMinute <= endMinute Then
    insideWindow = (nowMinute >= startMinute And nowMinute <= endMinute)
Else
    insideWindow = (nowMinute >= startMinute Or nowMinute <= endMinute)
End If
WScript.Echo "Current minute=" & nowMinute & " Maintenance window active=" & insideWindow
