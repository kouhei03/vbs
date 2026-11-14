' ファイル概要: 直近のシステムエラーをコンソールに表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 148_CriticalEventConsoleAlert.vbs <hours 1-24>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 直近のシステムエラーをコンソールに表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim hours, filter, service, events, eventItem, count
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 148_CriticalEventConsoleAlert.vbs <hours 1-24>"
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(0), hours) Then
    WScript.Echo "Hours must be an integer from 1 to 24."
    WScript.Quit 2
End If
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='System' AND Type='Error' AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    WScript.Echo "ALERT " & eventItem.TimeGenerated & " ID=" & eventItem.EventCode & _
        " Source=" & eventItem.SourceName
    count = count + 1
Next
WScript.Echo "System error events: " & count
WScript.Echo "Console alert only; no popup or event-log write was attempted."

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
