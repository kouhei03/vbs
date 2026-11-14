' ファイル概要: ローカルAdministratorsグループのメンバーを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 178_AdministratorsMembershipPreview.vbs
' 実行結果: ローカルAdministratorsグループのメンバーを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, groups, group, members, member, domainName, memberName
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Group を照会し、Name, SID を取得します。
Set groups = service.ExecQuery( _
    "SELECT Name, SID FROM Win32_Group WHERE SID='S-1-5-32-544'")
For Each group In groups
    WScript.Echo "Local Administrators group: " & group.Name
    Set members = group.Associators_("Win32_GroupUser")
    For Each member In members
        domainName = ""
        memberName = ""
        If Not IsNull(member.Domain) Then domainName = CStr(member.Domain)
        If Not IsNull(member.Name) Then memberName = CStr(member.Name)
        If Len(domainName) > 0 Then memberName = domainName & "\" & memberName
        WScript.Echo memberName
    Next
Next
WScript.Echo "Membership was inspected only; no user was added."
