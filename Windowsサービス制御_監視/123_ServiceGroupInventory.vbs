' サービスごとのServiceGroupレジストリ値を読み取ります。レジストリは変更しません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: ローカルWMI、サービス一覧、レジストリ照会に使う変数を宣言します。
Dim service, services, item, shell, registryPath, groupName
' 解説: ローカルWMI名前空間に接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス名と状態を読み取り専用で取得します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State FROM Win32_Service")
' 解説: ローカルレジストリを読み取るためのWScript.Shellオブジェクトを作成します。
Set shell = CreateObject("WScript.Shell")
' 解説: 各サービスのレジストリにあるグループ値を順番に調べます。
For Each item In services
    ' 解説: 対象サービスのGroup値があるレジストリパスを組み立てます。
    registryPath = "HKLM\SYSTEM\CurrentControlSet\Services\" & item.Name & "\Group"
    ' 解説: Group値が未設定の場合も他のサービスの照会を続けられるよう、次の1行だけエラーを確認します。
    On Error Resume Next
    ' 解説: サービスグループ値を読み取り専用で取得します。
    groupName = shell.RegRead(registryPath)
    ' 解説: 値が存在しない、またはアクセスできない場合はその状態を記録します。
    If Err.Number <> 0 Then
        ' 解説: 値が未設定または読み取れないことを表示用に設定します。
        groupName = "(not set or unreadable)"
        ' 解説: 確認済みのレジストリエラー状態をクリアします。
        Err.Clear
    ' 解説: エラー状態の分岐を終了します。
    End If
    ' 解説: 通常のエラー処理へ戻します。
    On Error GoTo 0
    ' 解説: サービス名、表示名、グループ値、状態を出力します。
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & groupName & vbTab & item.State
' 解説: 次のサービスの処理へ進みます。
Next
