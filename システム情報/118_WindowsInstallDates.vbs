' 解説: WMIのOSインストール日時とレジストリのセットアップ情報を読み取ります。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: service, systems, system, registry, result, installDate
Dim service, systems, system, registry, result, installDate
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set systems = service.ExecQuery("SELECT InstallDate FROM Win32_OperatingSystem")
' 解説: コレクション内の各項目を順に処理します: For Each system In systems
For Each system In systems
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "InstallDate=" & system.InstallDate
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 取得値または計算結果を変数へ格納します。
result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "InstallDate", installDate)
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: result = 0
If result = 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Setup InstallDate (Unix seconds)=" & installDate
' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
Else
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Setup InstallDate unavailable; WMI code=" & result
' 解説: 条件分岐ブロックを終了します。
End If
