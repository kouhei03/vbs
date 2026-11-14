' ファイル概要: DesktopInteract設定の読み取り。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 115_ServiceDesktopInteractStatus.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: DesktopInteract設定の読み取り処理の結果やエラーをコンソールに表示します。

' ローカルサービスのDesktopInteract設定を照会します。
Option Explicit
Dim serviceName, validator
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 115_ServiceDesktopInteractStatus.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Use a service name containing only letters, digits, dot, underscore, or hyphen."
    WScript.Quit 2
End If
Dim service, services, item, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DesktopInteract を取得します。
Set services = service.ExecQuery("SELECT Name, DesktopInteract FROM Win32_Service")
found = False
For Each item In services
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        WScript.Echo "Name=" & item.Name
        WScript.Echo "DesktopInteract=" & item.DesktopInteract
        WScript.Echo "ExitCode=" & item.ExitCode
        found = True
        Exit For
    End If
Next
If Not found Then
    WScript.Echo "Service not found: " & serviceName
    WScript.Quit 3
End If
