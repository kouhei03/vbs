' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、services、item）。主な用途: WMIサービス接続、サービス一覧、現在処理中のWMI項目。
Dim service, services, item
' 解説: このスクリプトで使う変数を宣言します（matched、upperName）。
Dim matched, upperName
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、Name, DisplayName, Stateを取得します。
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State FROM Win32_Service")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: matched = False
matched = False
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each item In services
For Each item In services
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: upperName = UCase(item.Name)
    upperName = UCase(item.Name)
    ' 解説: 条件「upperName = "W3SVC" Or upperName = "WAS" Or upperName = "MSSQLSERVER" Or upperName = "SQLSERVERAGENT" Or Left(upperName, 6) = "MSSQL$" Or Left(upperName, 9) = "SQLAGENT$" Then」を評価し、成立した場合に続く分岐処理を実行します。
    If upperName = "W3SVC" Or upperName = "WAS" Or _
        upperName = "MSSQLSERVER" Or upperName = "SQLSERVERAGENT" Or _
        Left(upperName, 6) = "MSSQL$" Or Left(upperName, 9) = "SQLAGENT$" Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: matched = True
        matched = True
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: item.Name & vbTab & item.DisplayName & vbTab & item.State
        WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 条件「Not matched Then」を評価し、成立した場合に続く分岐処理を実行します。
If Not matched Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "No matching SQL Server or IIS services were found."
    WScript.Echo "No matching SQL Server or IIS services were found."
' 解説: 条件分岐ブロックを終了します。
End If