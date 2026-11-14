' ファイル概要: サービスグループ情報の一覧。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 123_ServiceGroupInventory.vbs
' 実行結果: サービスグループ情報の一覧処理の結果やエラーをコンソールに表示します。

' サービスごとのServiceGroupレジストリ値を読み取ります。レジストリは変更しません。
Option Explicit
Dim service, services, item, shell, registryPath, groupName
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State を取得します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State FROM Win32_Service")
Set shell = CreateObject("WScript.Shell")
For Each item In services
    registryPath = "HKLM\SYSTEM\CurrentControlSet\Services\" & item.Name & "\Group"
    On Error Resume Next
    ' [処理解説] レジストリの現在値を読み取ります。設定値は変更しません。
    groupName = shell.RegRead(registryPath)
    If Err.Number <> 0 Then
        groupName = "(not set or unreadable)"
        Err.Clear
    End If
    On Error GoTo 0
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & groupName & vbTab & item.State
Next
