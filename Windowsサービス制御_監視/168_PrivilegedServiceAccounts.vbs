' 組み込みサービスアカウントで実行されるサービス名を照会します。設定変更はしません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: service, services, item, accountName
Dim service, services, item, accountName
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name, StartName, State FROM Win32_Service")
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each item In services
    ' 解説: Nullのログオンアカウントを文字列変換する前に除外します。
    If Not IsNull(item.StartName) Then
        ' 解説: アカウント名を大文字小文字の違いを無視して比較できる形にします。
        accountName = LCase(CStr(item.StartName))
        ' 解説: 組み込みサービスアカウントだけを選択して表示します。
        If accountName = "localsystem" Or accountName = "nt authority\system" Or accountName = "nt authority\localservice" Or accountName = "nt authority\networkservice" Or accountName = "localservice" Or accountName = "networkservice" Then
            ' 解説: 該当するサービス名、アカウント名、状態を表示します。
            WScript.Echo item.Name & vbTab & item.StartName & vbTab & item.State
        ' 解説: アカウントの選択条件を終了します。
        End If
    ' 解説: 条件分岐を終了します。
    End If
' 解説: 次の項目または反復へ進みます。
Next
