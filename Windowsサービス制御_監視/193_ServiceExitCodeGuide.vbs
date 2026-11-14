' サービスの終了コードを取得し、よくあるWin32コードの説明を表示します。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: serviceName, validator, service, services, item, found, codeValue, description
Dim serviceName, validator, service, services, item, found, codeValue, description
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 193_ServiceExitCodeGuide.vbs <service-name>"
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
    WScript.Echo "Invalid service name."
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name, State, ExitCode, ServiceSpecificExitCode FROM Win32_Service")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
found = False
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each item In services
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        ' 解説: 入力値、取得値、または計算結果を変数に格納します。
        codeValue = CLng(item.ExitCode)
        ' 解説: 入力値、取得値、または計算結果を変数に格納します。
        description = "Unmapped Win32 service exit code"
        ' 解説: 値に応じた複数分岐を開始します。
        Select Case codeValue
            ' 解説: 指定値に一致した場合の処理です。
            Case 0: description = "The operation completed successfully."
            ' 解説: 指定値に一致した場合の処理です。
            Case 2: description = "The system cannot find the file specified."
            ' 解説: 指定値に一致した場合の処理です。
            Case 5: description = "Access is denied."
            ' 解説: 指定値に一致した場合の処理です。
            Case 1053: description = "The service did not respond to the start or control request in time."
            ' 解説: 指定値に一致した場合の処理です。
            Case 1067: description = "The process terminated unexpectedly."
            ' 解説: 指定値に一致した場合の処理です。
            Case 1068: description = "A dependency service or group failed to start."
        ' 解説: Select Case分岐を終了します。
        End Select
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "State=" & item.State & " ExitCode=" & codeValue & " Meaning=" & description
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "ServiceSpecificExitCode=" & item.ServiceSpecificExitCode
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
    WScript.Echo "Service not found."
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 3
' 解説: 条件分岐を終了します。
End If
