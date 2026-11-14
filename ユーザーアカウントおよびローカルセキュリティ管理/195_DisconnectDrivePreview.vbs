' ファイル概要: 指定ネットワークドライブを確認する（切断なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 195_DisconnectDrivePreview.vbs <drive-letter>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定ネットワークドライブを確認する（切断なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim driveName, service, drives, drive, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 195_DisconnectDrivePreview.vbs <drive-letter>"
    WScript.Quit 2
End If
driveName = UCase(WScript.Arguments(0))
If Len(driveName) <> 2 Or Right(driveName, 1) <> ":" Or _
   InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", Left(driveName, 1), vbBinaryCompare) = 0 Then
    WScript.Echo "Drive must be a letter followed by a colon, such as Z:."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_LogicalDisk を照会し、DeviceID, ProviderName を取得します。
Set drives = service.ExecQuery( _
    "SELECT DeviceID, ProviderName FROM Win32_LogicalDisk WHERE DriveType=4")
found = False
For Each drive In drives
    If StrComp(drive.DeviceID, driveName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "Mapped drive: " & drive.DeviceID
        WScript.Echo "Provider: " & drive.ProviderName
    End If
Next
If Not found Then WScript.Echo "Mapped drive was not found."
WScript.Echo "Preview only; no network drive was disconnected."
