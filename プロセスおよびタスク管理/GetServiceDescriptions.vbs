' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、services、item、description）。主な用途: WMIサービス接続、サービス一覧、現在処理中のWMI項目。
Dim service, services, item, description
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、Name, DisplayName, Descriptionを取得します。
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, Description FROM Win32_Service")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Service name" & vbTab & "Display name" & vbTab & "Description"
WScript.Echo "Service name" & vbTab & "Display name" & vbTab & "Description"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each item In services
For Each item In services
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: description = item.Description
    description = item.Description
    ' 解説: 条件「IsNull(description) Then」を評価し、成立した場合に続く分岐処理を実行します。
    If IsNull(description) Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: description = "(no description)"
        description = "(no description)"
    ' 解説: 先行する条件が不成立の場合に評価する追加条件です: Len(description) = 0 Then
    ElseIf Len(description) = 0 Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: description = "(no description)"
        description = "(no description)"
    ' 解説: 条件分岐ブロックを終了します。
    End If
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: item.Name & vbTab & item.DisplayName & vbTab & description
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & description
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
