' 概要: 指定したドライブ文字が接続されているかだけ確認します。切断は行いません。
' 実行方法: Usage: cscript //nologo 002_NetworkDriveDisconnectPreview.vbs <drive-letter>
' 処理結果: 指定ドライブの接続状態確認（切断なし）。処理結果や状態はコンソールに表示されます。
' コード解説: スクリプト引数を使って対象を特定し、検証後に処理します。
Option Explicit
Dim letter, network, drives, i, found
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 002_NetworkDriveDisconnectPreview.vbs <drive-letter>"
    WScript.Quit 2
End If
' 主処理: 指定ドライブの接続状態確認（切断なし）。
letter = UCase(Replace(WScript.Arguments(0), ":", "")) & ":"
If Len(letter) <> 2 Or Not (Left(letter, 1) >= "A" And Left(letter, 1) <= "Z") Then WScript.Quit 2
Set network = CreateObject("WScript.Network")
Set drives = network.EnumNetworkDrives
found = False
For i = 0 To drives.Count - 1 Step 2
    If UCase(drives.Item(i)) = letter Then
        WScript.Echo "Connected: " & letter & " -> " & drives.Item(i + 1)
        found = True
    End If
Next
If Not found Then WScript.Echo "No network mapping found for " & letter
WScript.Echo "Preview only; no connection was removed."
