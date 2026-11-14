' ファイル概要: ログオン・ログオフのイベントIDと時刻を表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 153_LogonLogoffTimeline.vbs <hours 1-168>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ログオン・ログオフのイベントIDと時刻を表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim hours, filter, service, events, eventItem
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 153_LogonLogoffTimeline.vbs <hours 1-168>"
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(0), hours) Then
    WScript.Echo "Hours must be an integer from 1 to 168."
    WScript.Quit 2
End If
Set filter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
filter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, TimeGenerated を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Security' AND (EventCode=4624 OR EventCode=4634 OR EventCode=4647) " & _
    "AND TimeGenerated >= '" & filter.Value & "'")
WScript.Echo "Time" & vbTab & "Event ID"
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode
Next
WScript.Echo "Account and network details are omitted."

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
