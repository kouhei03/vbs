' ファイル概要: 直近のシステムエラーをコンソールに表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 148_CriticalEventConsoleAlert.vbs <hours 1-24>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 直近のシステムエラーをコンソールに表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim hours, dateFilter, service, events, eventItem, count
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 148_CriticalEventConsoleAlert.vbs <hours 1-24>"
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(0), hours) Then
    WScript.Echo "Hours must be an integer from 1 to 24."
    WScript.Quit 2
End If
Set dateFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
dateFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, SourceName, TimeGenerated を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='System' AND Type='Error' AND TimeGenerated >= '" & dateFilter.Value & "'")
count = 0
For Each eventItem In events
    WScript.Echo "ALERT " & eventItem.TimeGenerated & " ID=" & eventItem.EventCode & _
        " Source=" & eventItem.SourceName
    count = count + 1
Next
WScript.Echo "System error events: " & count
WScript.Echo "Console alert only; no popup or event-log write was attempted."

' [処理解説] 数字以外の入力を除外してから整数へ変換し、指定範囲に収まるか判定します。
Function ParseHours(text, value)
    Dim i, character
    ParseHours = False
    If Len(text) = 0 Or Len(text) > 2 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseHours = (value >= 1 And value <= 24)
End Function
