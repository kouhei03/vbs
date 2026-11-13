' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、profiles、profile）。主な用途: WMIサービス接続、プロファイル一覧、ユーザープロファイル。
Dim service, profiles, profile
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\standardcimv2")
' 解説: WMIのMSFT_NetFirewallProfile（MSFT_NetFirewallProfile）を読み取り専用で照会し、Name, Enabledを取得します。
Set profiles = service.ExecQuery( _
    "SELECT Name, Enabled FROM MSFT_NetFirewallProfile")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Firewall profile" & vbTab & "Enabled"
WScript.Echo "Firewall profile" & vbTab & "Enabled"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each profile In profiles
For Each profile In profiles
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: profile.Name & vbTab & profile.Enabled
    WScript.Echo profile.Name & vbTab & profile.Enabled
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
