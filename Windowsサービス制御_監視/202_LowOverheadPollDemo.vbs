' ファイル概要: 低頻度・回数限定のサービス照会。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 202_LowOverheadPollDemo.vbs
' 実行結果: 低頻度・回数限定のサービス照会処理の結果やエラーをコンソールに表示します。

' 指定間隔で限定回数だけサービス数を照会し、簡単な低頻度ポーリング例を示します。
Option Explicit
Dim service, services, sample
Set service = GetObject("winmgmts:\\.\root\cimv2")
For sample = 1 To 3
    Set services = service.ExecQuery("SELECT Name FROM Win32_Service")
    WScript.Echo CStr(Now) & vbTab & "Sample=" & sample & vbTab & "Services=" & services.Count
    If sample < 3 Then WScript.Sleep 10000
Next
