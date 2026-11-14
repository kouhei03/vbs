' ローカルサービス構成の基本情報を読み取ります。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: serviceName, validator
Dim serviceName, validator
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 120_ServiceConfigurationReadOnly.vbs <service-name>"
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: 入力値、取得値、または計算結果を変数に格納します。
serviceName = Trim(WScript.Arguments(0))
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 入力値、取得値、または計算結果を変数に格納します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Not validator.Test(serviceName) Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Use a service name containing only letters, digits, dot, underscore, or hyphen."
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: このスクリプトで使用する変数を宣言します: service, services, item, found
Dim service, services, item, found
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode, StartName, PathName, Description, ExitCode FROM Win32_Service")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
found = False
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each item In services
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "Name=" & item.Name
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "DisplayName=" & item.DisplayName
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "State=" & item.State
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "StartMode=" & item.StartMode
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "Account=" & item.StartName
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "ImagePath=" & item.PathName
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "Description=" & item.Description
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "ExitCode=" & item.ExitCode
        ' 解説: 入力値、取得値、または計算結果を変数に格納します。
        found = True
        ' 解説: ローカルの読み取り専用診断処理を進めます。
        Exit For
    ' 解説: 条件分岐を終了します。
    End If
' 解説: 次の項目または反復へ進みます。
Next
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Not found Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Service not found: " & serviceName
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 3
' 解説: 条件分岐を終了します。
End If
