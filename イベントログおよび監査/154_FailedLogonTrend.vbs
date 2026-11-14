' ファイル概要: ログオン失敗の件数を集計する（アカウント名とアドレスは非表示）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 154_FailedLogonTrend.vbs <hours 1-168>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ログオン失敗の件数を集計する（アカウント名とアドレスは非表示）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim hours, filter, service, events, eventItem, count
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 154_FailedLogonTrend.vbs <hours 1-168>"
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(0), hours) Then
    WScript.Echo "Hours must be an integer from 1 to 168."
    WScript.Quit 2
End If
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='Security' " & _
    "AND EventCode=4625 AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    count = count + 1
Next
WScript.Echo "Failed logon count (Event ID 4625): " & count
WScript.Echo "Account names, addresses, and event messages were not collected."

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
