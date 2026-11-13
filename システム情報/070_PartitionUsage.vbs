' 解説: 論理ドライブごとの総量、使用量、空き容量を読み取ります。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: service, drives, drive, usedBytes, usedPercent
Dim service, drives, drive, usedBytes, usedPercent
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set drives = service.ExecQuery("SELECT DeviceID, Size, FreeSpace, FileSystem FROM Win32_LogicalDisk WHERE DriveType = 3")
' 解説: コレクション内の各項目を順に処理します: For Each drive In drives
For Each drive In drives
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Not IsNull(drive.Size) And Not IsNull(drive.FreeSpace)
    If Not IsNull(drive.Size) And Not IsNull(drive.FreeSpace) Then
        ' 解説: 取得値または計算結果を変数へ格納します。
        usedBytes = CDbl(drive.Size) - CDbl(drive.FreeSpace)
        ' 解説: 取得値または計算結果を変数へ格納します。
        usedPercent = 0
        ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: CDbl(drive.Size) > 0
        If CDbl(drive.Size) > 0 Then usedPercent = (usedBytes / CDbl(drive.Size)) * 100
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo drive.DeviceID & " FileSystem=" & drive.FileSystem
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo "Total bytes=" & drive.Size & " Used bytes=" & usedBytes & " Free bytes=" & drive.FreeSpace
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo "Used percent=" & FormatNumber(usedPercent, 2)
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
