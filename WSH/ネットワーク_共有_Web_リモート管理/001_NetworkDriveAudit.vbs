' 現在割り当てられているネットワークドライブを表示します。接続・切断はしません。
Option Explicit
Dim network, drives, i
Set network = CreateObject("WScript.Network")
Set drives = network.EnumNetworkDrives
For i = 0 To drives.Count - 1 Step 2
    WScript.Echo drives.Item(i) & vbTab & drives.Item(i + 1)
Next
