' 概要: 現在割り当てられているネットワークドライブを表示します。接続・切断はしません。
' 実行方法: cscript //nologo 001_NetworkDriveAudit.vbs
' 処理結果: 現在のネットワークドライブ一覧（接続変更なし）。処理結果や状態はコンソールに表示されます。
' コード解説: Windowsまたはネットワークの状態を確認し、結果を表示します。
Option Explicit
Dim network, drives, i
' 主処理: 現在のネットワークドライブ一覧（接続変更なし）。
Set network = CreateObject("WScript.Network")
Set drives = network.EnumNetworkDrives
For i = 0 To drives.Count - 1 Step 2
    WScript.Echo drives.Item(i) & vbTab & drives.Item(i + 1)
Next
