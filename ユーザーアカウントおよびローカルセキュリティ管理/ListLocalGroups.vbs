' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、groups、group）。主な用途: WMIサービス接続、グループ一覧、グループ情報。
Dim service, groups, group
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのローカルグループ（Win32_Group）を読み取り専用で照会し、Name, Descriptionを取得します。 固定絞り込み条件: LocalAccount=True。
Set groups = service.ExecQuery( _
    "SELECT Name, Description FROM Win32_Group WHERE LocalAccount=True")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Local group" & vbTab & "Description"
WScript.Echo "Local group" & vbTab & "Description"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each group In groups
For Each group In groups
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: group.Name & vbTab & group.Description
    WScript.Echo group.Name & vbTab & group.Description
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
