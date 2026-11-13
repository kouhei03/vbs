' 解説: 開始/終了を分単位で受け取り、現在時刻が保守時間帯か判定します。プロセスは停止しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: startMinute, endMinute, nowMinute, insideWindow
Dim startMinute, endMinute, nowMinute, insideWindow
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 2
If WScript.Arguments.Count <> 2 Then WScript.Echo "Usage: cscript //nologo 090_MaintenanceWindowCheck.vbs <start-minute-0-1439> <end-minute-0-1439>": WScript.Quit 2
' 解説: 条件を判定し、必要な分岐を選択します: If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1))
If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Echo "Times must be minute numbers.": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
startMinute = CLng(WScript.Arguments(0))
' 解説: 取得値または計算結果を変数に格納します。
endMinute = CLng(WScript.Arguments(1))
' 解説: 条件を判定し、必要な分岐を選択します: If startMinute < 0 Or startMinute > 1439 Or endMinute < 0 Or endMinute > 1439
If startMinute < 0 Or startMinute > 1439 Or endMinute < 0 Or endMinute > 1439 Then WScript.Echo "Minutes must be between 0 and 1439.": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
nowMinute = Hour(Now) * 60 + Minute(Now)
' 解説: 条件を判定し、必要な分岐を選択します: If startMinute <= endMinute
If startMinute <= endMinute Then
    ' 解説: 取得値または計算結果を変数に格納します。
    insideWindow = (nowMinute >= startMinute And nowMinute <= endMinute)
' 解説: 先行条件が成立しない場合の代替処理へ進みます。
Else
    ' 解説: 取得値または計算結果を変数に格納します。
    insideWindow = (nowMinute >= startMinute Or nowMinute <= endMinute)
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Current minute=" & nowMinute & " Maintenance window active=" & insideWindow
