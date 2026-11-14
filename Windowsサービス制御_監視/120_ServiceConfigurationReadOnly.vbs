' ファイル概要: 基本サービス構成情報の読み取り。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 120_ServiceConfigurationReadOnly.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 基本サービス構成情報の読み取り処理の結果やエラーをコンソールに表示します。

' ローカルサービス構成の基本情報を読み取ります。
Option Explicit
Dim serviceName, validator
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 120_ServiceConfigurationReadOnly.vbs <service-name>"
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
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State, StartMode, StartName, PathName, Description, ExitCode を取得します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode, StartName, PathName, Description, ExitCode FROM Win32_Service")
found = False
For Each item In services
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        WScript.Echo "Name=" & item.Name
        WScript.Echo "DisplayName=" & item.DisplayName
        WScript.Echo "State=" & item.State
        WScript.Echo "StartMode=" & item.StartMode
        WScript.Echo "Account=" & item.StartName
        WScript.Echo "ImagePath=" & item.PathName
        WScript.Echo "Description=" & item.Description
        WScript.Echo "ExitCode=" & item.ExitCode
        found = True
        Exit For
    End If
Next
If Not found Then
    WScript.Echo "Service not found: " & serviceName
    WScript.Quit 3
End If
