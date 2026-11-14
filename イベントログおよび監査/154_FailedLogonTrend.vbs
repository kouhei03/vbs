' ファイル概要: ログオン失敗の件数を集計する（アカウント名とアドレスは非表示）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 154_FailedLogonTrend.vbs <hours 1-168>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ログオン失敗の件数を集計する（アカウント名とアドレスは非表示）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim hours, dateFilter, service, events, eventItem, count
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 154_FailedLogonTrend.vbs <hours 1-168>"
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(0), hours) Then
    WScript.Echo "Hours must be an integer from 1 to 168."
    WScript.Quit 2
End If
Set dateFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
dateFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='Security' " & _
    "AND EventCode=4625 AND TimeGenerated >= '" & dateFilter.Value & "'")
count = 0
For Each eventItem In events
    count = count + 1
Next
WScript.Echo "Failed logon count (Event ID 4625): " & count
WScript.Echo "Account names, addresses, and event messages were not collected."

' [処理解説] 数字以外の入力を除外してから整数へ変換し、指定範囲に収まるか判定します。
Function ParseHours(text, value)
    Dim i, character
    ParseHours = False
    If Len(text) = 0 Or Len(text) > 3 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseHours = (value >= 1 And value <= 168)
End Function
