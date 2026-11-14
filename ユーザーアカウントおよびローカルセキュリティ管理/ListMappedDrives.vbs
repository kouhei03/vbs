' ファイル概要: 割り当て済みネットワークドライブを一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ListMappedDrives.vbs
' 実行結果: 割り当て済みネットワークドライブを一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim network, mappings, index
Set network = CreateObject("WScript.Network")
Set mappings = network.EnumNetworkDrives

If mappings.Count = 0 Then
    WScript.Echo "No network drives are mapped."
Else
    WScript.Echo "Drive" & vbTab & "Network path"
    ' [処理解説] EnumNetworkDrivesの結果はドライブ名とネットワークパスの対なので、2項目ずつ処理します。
    For index = 0 To mappings.Count - 1 Step 2
        WScript.Echo mappings.Item(index) & vbTab & mappings.Item(index + 1)
    Next
End If
