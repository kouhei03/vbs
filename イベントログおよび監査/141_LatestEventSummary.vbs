' ファイル概要: 過去24時間のイベントメタデータを最大20件表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 141_LatestEventSummary.vbs <System|Application|Security>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 過去24時間のイベントメタデータを最大20件表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim logName, filter, service, events, eventItem, count
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 141_LatestEventSummary.vbs <System|Application|Security>"
    WScript.Quit 2
End If
logName = WScript.Arguments(0)
If Not IsAllowedLog(logName) Then
    WScript.Echo "Unsupported event log."
    WScript.Quit 2
End If

Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("d", -1, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM Win32_NTLogEvent " & _
    "WHERE Logfile='" & logName & "' AND TimeGenerated >= '" & filter.Value & "'")
WScript.Echo "Recent entries in " & logName & " (last 24 hours, metadata only):"
count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & vbTab & _
        eventItem.SourceName & vbTab & eventItem.RecordNumber
    count = count + 1
    If count >= 20 Then Exit For
Next
If count = 0 Then WScript.Echo "No entries returned."

Function IsAllowedLog(name)
    IsAllowedLog = (StrComp(name, "System", vbTextCompare) = 0 Or _
        StrComp(name, "Application", vbTextCompare) = 0 Or _
        StrComp(name, "Security", vbTextCompare) = 0)
End Function
