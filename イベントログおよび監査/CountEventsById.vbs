' ファイル概要: 指定イベントIDの件数を集計する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo CountEventsById.vbs <System|Application|Security> <event-id> <hours 1-168>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定イベントIDの件数を集計する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim logName, eventId, hours, timeFilter, service, events, eventItem, count
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo CountEventsById.vbs <System|Application|Security> <event-id> <hours 1-168>"
    WScript.Quit 2
End If

logName = WScript.Arguments(0)
If Not IsAllowedLog(logName) Then
    WScript.Echo "Log must be System, Application, or Security."
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(1), 0, 65535, eventId) Then
    WScript.Echo "Event ID must be an integer from 0 to 65535."
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(2), 1, 168, hours) Then
    WScript.Echo "Hours must be a whole number from 1 to 168."
    WScript.Quit 2
End If

Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='" & logName & _
    "' AND EventCode=" & eventId & _
    " AND TimeGenerated >= '" & timeFilter.Value & "'")

count = 0
For Each eventItem In events
    count = count + 1
Next
WScript.Echo "Event " & eventId & " in " & logName & _
    " during the last " & hours & " hours: " & count

Function IsAllowedLog(name)
    IsAllowedLog = (StrComp(name, "System", vbTextCompare) = 0 Or _
        StrComp(name, "Application", vbTextCompare) = 0 Or _
        StrComp(name, "Security", vbTextCompare) = 0)
End Function

Function ParseInteger(text, minimum, maximum, value)
    Dim i, ch
    ParseInteger = False
    If Len(text) = 0 Or Len(text) > 5 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If InStr(1, "0123456789", ch, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseInteger = (value >= minimum And value <= maximum)
End Function
