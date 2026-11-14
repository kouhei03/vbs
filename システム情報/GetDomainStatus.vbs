' ファイル概要: ドメインまたはワークグループへの参加状態を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetDomainStatus.vbs
' 実行結果: ドメインまたはワークグループへの参加状態を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, systems, system
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery( _
    "SELECT PartOfDomain, Domain, Workgroup FROM Win32_ComputerSystem")

For Each system In systems
    If system.PartOfDomain Then
        WScript.Echo "Join type: Domain"
        WScript.Echo "Domain: " & system.Domain
    Else
        WScript.Echo "Join type: Workgroup"
        WScript.Echo "Workgroup: " & system.Workgroup
    End If
Next
