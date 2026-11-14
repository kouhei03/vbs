' ファイル概要: 指定サービスのログオンアカウントを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetServiceLogonAccount.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定サービスのログオンアカウントを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item, serviceName, found, accountName
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo GetServiceLogonAccount.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
If Len(serviceName) = 0 Then
    WScript.Echo "Service name must not be empty."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, StartName を取得します。
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, StartName FROM Win32_Service")
found = False
For Each item In services
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        found = True
        accountName = item.StartName
        If IsNull(accountName) Then
            accountName = "(not reported)"
        ElseIf Len(accountName) = 0 Then
            accountName = "(not reported)"
        End If
        WScript.Echo item.DisplayName & " runs as: " & accountName
        Exit For
    End If
Next
If Not found Then
    WScript.Echo "Service not found: " & serviceName
    WScript.Quit 3
End If
