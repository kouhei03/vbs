' 概要: ローカルコンピューターのドメイン参加状態を表示します。信頼関係の修復はしません。
' 実行方法: cscript //nologo 039_DomainMembershipAudit.vbs
' 処理結果: ローカルのドメイン参加状態確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, systems, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set systems = service.ExecQuery("SELECT Name, PartOfDomain, Domain FROM Win32_ComputerSystem")
For Each item In systems
    WScript.Echo "Computer=" & item.Name & " PartOfDomain=" & item.PartOfDomain & " Domain=" & item.Domain
Next
WScript.Echo "Read-only membership check; no secure-channel repair was attempted."
