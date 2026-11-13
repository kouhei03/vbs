' 解説: VPN/トンネル名を含むローカルネットワークアダプターの接続状態を表示します。VPNは接続/切断しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, adapters, adapter, label
Dim service, adapters, adapter, label
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set adapters = service.ExecQuery("SELECT Name,NetConnectionStatus,NetEnabled,PhysicalAdapter FROM Win32_NetworkAdapter")
' 解説: コレクション内の各項目を順番に処理します: For Each adapter In adapters
For Each adapter In adapters
    ' 解説: 取得値または計算結果を変数に格納します。
    label = LCase(CStr(adapter.Name))
    ' 解説: 条件を判定し、必要な分岐を選択します: If InStr(label,"vpn")>0 Or InStr(label,"tunnel")>0 Or InStr(label,"wireguard")>0
    If InStr(label,"vpn")>0 Or InStr(label,"tunnel")>0 Or InStr(label,"wireguard")>0 Then
        ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
        WScript.Echo adapter.Name & " Status=" & adapter.NetConnectionStatus & " Enabled=" & adapter.NetEnabled
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
