' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、logs、logItem）。主な用途: WMIサービス接続。
Dim service, logs, logItem
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのイベントログファイル設定（Win32_NTEventLogFile）を読み取り専用で照会し、Name, FileName, FileSize, MaxFileSize, OverwriteOutdated, OverwritePolicyを取得します。
Set logs = service.ExecQuery( _
    "SELECT Name, FileName, FileSize, MaxFileSize, " & _
    "OverwriteOutdated, OverwritePolicy FROM Win32_NTEventLogFile")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Log" & vbTab & "Current size (bytes)" & vbTab & "Maximum size (bytes)" & vbTab & "Overwrite policy"
WScript.Echo "Log" & vbTab & "Current size (bytes)" & vbTab & _
    "Maximum size (bytes)" & vbTab & "Overwrite policy"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each logItem In logs
For Each logItem In logs
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: logItem.Name & vbTab & logItem.FileSize & vbTab & logItem.MaxFileSize & vbTab & logItem.OverwritePolicy
    WScript.Echo logItem.Name & vbTab & logItem.FileSize & vbTab & _
        logItem.MaxFileSize & vbTab & logItem.OverwritePolicy
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
