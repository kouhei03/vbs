' ファイル概要: バックアップ/VSS関連サービスの状態確認。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 153_BackupVssStatus.vbs
' 実行結果: バックアップ/VSS関連サービスの状態確認処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, services, item, wantedNames, wantedName, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State, StartMode, ExitCode を取得します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode, ExitCode FROM Win32_Service")
wantedNames = Array("VSS", "swprv")
For Each wantedName In wantedNames
    found = False
    For Each item In services
        If StrComp(item.Name, wantedName, vbTextCompare) = 0 Then
            WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State & vbTab & item.StartMode & vbTab & item.ExitCode
            found = True
            Exit For
        End If
    Next
    If Not found Then WScript.Echo wantedName & vbTab & "Not installed or not visible"
Next
