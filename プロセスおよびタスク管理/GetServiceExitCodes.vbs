' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、services、item）。主な用途: WMIサービス接続、サービス一覧、現在処理中のWMI項目。
Dim service, services, item
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、Name, State, ExitCode, ServiceSpecificExitCodeを取得します。
Set services = service.ExecQuery( _
    "SELECT Name, State, ExitCode, ServiceSpecificExitCode FROM Win32_Service")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Service name" & vbTab & "State" & vbTab & "Exit code" & vbTab & "Service-specific exit code"
WScript.Echo "Service name" & vbTab & "State" & vbTab & _
    "Exit code" & vbTab & "Service-specific exit code"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each item In services
For Each item In services
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: item.Name & vbTab & item.State & vbTab & item.ExitCode & vbTab & item.ServiceSpecificExitCode
    WScript.Echo item.Name & vbTab & item.State & vbTab & _
        item.ExitCode & vbTab & item.ServiceSpecificExitCode
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
