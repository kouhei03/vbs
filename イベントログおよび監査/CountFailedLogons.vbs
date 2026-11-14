' ファイル概要: 指定期間のログオン失敗イベントを集計する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo CountFailedLogons.vbs <hours 1-168>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定期間のログオン失敗イベントを集計する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim hours, timeFilter, service, events, eventItem, count

' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo CountFailedLogons.vbs <hours 1-168>"
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(0), hours) Then
    WScript.Echo "Hours must be a whole number from 1 to 168."
    WScript.Quit 2
End If

Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
timeFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='Security' " & _
    "AND EventCode=4625 AND TimeGenerated >= '" & timeFilter.Value & "'")

count = 0
For Each eventItem In events
    count = count + 1
Next
WScript.Echo "Failed logon events (4625) in the last " & hours & _
    " hours: " & count
WScript.Echo "Only a count is displayed; account and network details are omitted."

' [処理解説] 数字以外の入力を除外してから整数へ変換し、指定範囲に収まるか判定します。
Function ParseHours(text, value)
    Dim i, ch
    ParseHours = False
    If Len(text) = 0 Or Len(text) > 3 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If InStr(1, "0123456789", ch, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseHours = (value >= 1 And value <= 168)
End Function
