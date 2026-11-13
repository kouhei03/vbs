' 解説: 論理ドライブのDriveTypeコードを読み、代表的なドライブ種別に変換します。
Option Explicit
' 解説: WMI接続、論理ドライブ、および表示用種別名を保持します。
Dim service, drives, drive, driveTypeName
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set drives = service.ExecQuery("SELECT DeviceID, DriveType, ProviderName FROM Win32_LogicalDisk")
' 解説: コレクション内の各項目を順に処理します: For Each drive In drives
For Each drive In drives
    ' 解説: 値に応じて対応するケースを選択します: drive.DriveType
    Select Case drive.DriveType
        ' 解説: この値または条件に該当する場合の分岐です: Case 0: driveTypeName = "Unknown"
        Case 0: driveTypeName = "Unknown"
        ' 解説: この値または条件に該当する場合の分岐です: Case 1: driveTypeName = "No root directory"
        Case 1: driveTypeName = "No root directory"
        ' 解説: この値または条件に該当する場合の分岐です: Case 2: driveTypeName = "Removable"
        Case 2: driveTypeName = "Removable"
        ' 解説: この値または条件に該当する場合の分岐です: Case 3: driveTypeName = "Local fixed disk"
        Case 3: driveTypeName = "Local fixed disk"
        ' 解説: この値または条件に該当する場合の分岐です: Case 4: driveTypeName = "Network"
        Case 4: driveTypeName = "Network"
        ' 解説: この値または条件に該当する場合の分岐です: Case 5: driveTypeName = "CD-ROM"
        Case 5: driveTypeName = "CD-ROM"
        ' 解説: この値または条件に該当する場合の分岐です: Case 6: driveTypeName = "RAM disk"
        Case 6: driveTypeName = "RAM disk"
        ' 解説: この値または条件に該当する場合の分岐です: Case Else: driveTypeName = "Unrecognized type " & drive.DriveType
        Case Else: driveTypeName = "Unrecognized type " & drive.DriveType
    ' 解説: Select Caseの選択分岐を終了します。
    End Select
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo drive.DeviceID & ": " & driveTypeName & " Provider=" & drive.ProviderName
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
