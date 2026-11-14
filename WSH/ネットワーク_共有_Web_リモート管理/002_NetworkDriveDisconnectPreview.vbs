' 指定したドライブ文字が接続されているかだけ確認します。切断は行いません。
Option Explicit
Dim letter, network, drives, i, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 002_NetworkDriveDisconnectPreview.vbs <drive-letter>"
    WScript.Quit 2
End If
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
