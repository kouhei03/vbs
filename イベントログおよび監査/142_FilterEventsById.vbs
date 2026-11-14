' ファイル概要: 指定イベントIDと時間範囲でローカルログを検索する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 142_FilterEventsById.vbs <System|Application|Security> <event-id> <hours 1-168>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定イベントIDと時間範囲でローカルログを検索する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim logName, eventId, hours, service, events, eventItem, filter, count
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 142_FilterEventsById.vbs <System|Application|Security> <event-id> <hours 1-168>"
    WScript.Quit 2
End If
logName = WScript.Arguments(0)
If Not IsAllowedLog(logName) Then
    WScript.Echo "Unsupported event log."
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(1), 0, 65535, eventId) Then
    WScript.Echo "Event ID must be an integer from 0 to 65535."
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(2), 1, 168, hours) Then
    WScript.Echo "Hours must be an integer from 1 to 168."
    WScript.Quit 2
End If

Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM Win32_NTLogEvent " & _
    "WHERE Logfile='" & logName & "' AND EventCode=" & eventId & _
    " AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & vbTab & _
        eventItem.SourceName & vbTab & eventItem.RecordNumber
    count = count + 1
Next
WScript.Echo "Matching event count: " & count

Function IsAllowedLog(name)
    IsAllowedLog = (StrComp(name, "System", vbTextCompare) = 0 Or _
        StrComp(name, "Application", vbTextCompare) = 0 Or _
        StrComp(name, "Security", vbTextCompare) = 0)
End Function

Function ParseInteger(text, minimum, maximum, value)
    Dim i, character
    ParseInteger = False
    If Len(text) = 0 Or Len(text) > 5 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseInteger = (value >= minimum And value <= maximum)
End Function
