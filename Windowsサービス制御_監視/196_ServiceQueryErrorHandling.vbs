' ファイル概要: WMI照会エラー処理の例。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 196_ServiceQueryErrorHandling.vbs
' 実行結果: WMI照会エラー処理の例処理の結果やエラーをコンソールに表示します。

' WMI照会の失敗を明示的に報告する読み取り専用エラーハンドリング例です。
Option Explicit
Dim service, services, errorNumber, errorDescription
On Error Resume Next
Set service = GetObject("winmgmts:\\.\root\cimv2")
If Err.Number <> 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    WScript.Echo "WMI connection failed: " & errorNumber & " " & errorDescription
    WScript.Quit 1
End If
Set services = service.ExecQuery("SELECT Name FROM Win32_Service")
If Err.Number <> 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    WScript.Echo "Service query failed: " & errorNumber & " " & errorDescription
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "Service query succeeded; count=" & services.Count
