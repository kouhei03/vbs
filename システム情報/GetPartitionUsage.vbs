' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、drives、drive、usedBytes、usedPercent）。主な用途: WMIサービス接続、ドライブ一覧、ドライブ情報、使用容量（バイト）、使用率（%）。
Dim service, drives, drive, usedBytes, usedPercent
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの論理ドライブ（Win32_LogicalDisk）を読み取り専用で照会し、DeviceID, Size, FreeSpaceを取得します。 固定絞り込み条件: DriveType=3。
Set drives = service.ExecQuery( _
    "SELECT DeviceID, Size, FreeSpace FROM Win32_LogicalDisk WHERE DriveType=3")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each drive In drives
For Each drive In drives
    ' 解説: 条件「Not IsNull(drive.Size) Then」を評価し、成立した場合に続く分岐処理を実行します。
    If Not IsNull(drive.Size) Then
        ' 解説: 条件「Not IsNull(drive.FreeSpace) Then」を評価し、成立した場合に続く分岐処理を実行します。
        If Not IsNull(drive.FreeSpace) Then
            ' 解説: 条件「CDbl(drive.Size) > 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
            If CDbl(drive.Size) > 0 Then
                ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: usedBytes = CDbl(drive.Size) - CDbl(drive.FreeSpace)
                usedBytes = CDbl(drive.Size) - CDbl(drive.FreeSpace)
                ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: usedPercent = Round((usedBytes / CDbl(drive.Size)) * 100, 1)
                usedPercent = Round((usedBytes / CDbl(drive.Size)) * 100, 1)
                ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: drive.DeviceID & " used (bytes): " & usedBytes
                WScript.Echo drive.DeviceID & " used (bytes): " & usedBytes
                ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: drive.DeviceID & " free (bytes): " & drive.FreeSpace
                WScript.Echo drive.DeviceID & " free (bytes): " & drive.FreeSpace
                ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: drive.DeviceID & " usage (%): " & usedPercent
                WScript.Echo drive.DeviceID & " usage (%): " & usedPercent
            ' 解説: 条件分岐ブロックを終了します。
            End If
        ' 解説: 条件分岐ブロックを終了します。
        End If
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
