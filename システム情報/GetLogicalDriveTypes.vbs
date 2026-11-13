' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、drives、drive、driveType）。主な用途: WMIサービス接続、ドライブ一覧、ドライブ情報。
Dim service, drives, drive, driveType
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの論理ドライブ（Win32_LogicalDisk）を読み取り専用で照会し、DeviceID, DriveTypeを取得します。
Set drives = service.ExecQuery( _
    "SELECT DeviceID, DriveType FROM Win32_LogicalDisk")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each drive In drives
For Each drive In drives
    ' 解説: 値ごとに処理を分ける選択分岐を開始します: drive.DriveType
    Select Case drive.DriveType
        ' 解説: Select Caseの値が「0: driveType = "Unknown"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 0: driveType = "Unknown"
        ' 解説: Select Caseの値が「1: driveType = "No root directory"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 1: driveType = "No root directory"
        ' 解説: Select Caseの値が「2: driveType = "Removable"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 2: driveType = "Removable"
        ' 解説: Select Caseの値が「3: driveType = "Fixed"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 3: driveType = "Fixed"
        ' 解説: Select Caseの値が「4: driveType = "Network"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 4: driveType = "Network"
        ' 解説: Select Caseの値が「5: driveType = "CD-ROM"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 5: driveType = "CD-ROM"
        ' 解説: Select Caseの値が「6: driveType = "RAM disk"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 6: driveType = "RAM disk"
        ' 解説: Select CaseのどのCase値にも一致しない場合に実行する既定分岐です。
        Case Else: driveType = "Other"
    ' 解説: Select Caseの選択分岐を終了します。
    End Select
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: drive.DeviceID & ": " & driveType
    WScript.Echo drive.DeviceID & ": " & driveType
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
