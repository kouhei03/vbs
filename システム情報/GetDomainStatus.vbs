' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、systems、system）。主な用途: WMIサービス接続、OS情報一覧、OS情報。
Dim service, systems, system
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのコンピューター/ドメイン参加（Win32_ComputerSystem）を読み取り専用で照会し、PartOfDomain, Domain, Workgroupを取得します。
Set systems = service.ExecQuery( _
    "SELECT PartOfDomain, Domain, Workgroup FROM Win32_ComputerSystem")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each system In systems
For Each system In systems
    ' 解説: 条件「system.PartOfDomain Then」を評価し、成立した場合に続く分岐処理を実行します。
    If system.PartOfDomain Then
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Join type: Domain"
        WScript.Echo "Join type: Domain"
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Domain: " & system.Domain
        WScript.Echo "Domain: " & system.Domain
    ' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
    Else
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Join type: Workgroup"
        WScript.Echo "Join type: Workgroup"
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Workgroup: " & system.Workgroup
        WScript.Echo "Workgroup: " & system.Workgroup
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
