' 無効状態のサービスを列挙し、復旧検討用の情報を表示します。設定は変更しません。
Option Explicit
Dim service, services, item, count
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode FROM Win32_Service WHERE StartMode='Disabled'")
count = 0
For Each item In services
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State & vbTab & item.StartMode
    count = count + 1
Next
WScript.Echo "Disabled services found=" & count & "; no startup settings were changed."
