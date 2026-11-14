' ローカルコンピューターのドメイン参加状態を表示します。信頼関係の修復はしません。
Option Explicit
Dim service, systems, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT Name, PartOfDomain, Domain FROM Win32_ComputerSystem")
For Each item In systems
    WScript.Echo "Computer=" & item.Name & " PartOfDomain=" & item.PartOfDomain & " Domain=" & item.Domain
Next
WScript.Echo "Read-only membership check; no secure-channel repair was attempted."
