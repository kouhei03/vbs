' ファイル概要: 状態・スタートアップ種別の集計。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 205_ServiceInventorySummary.vbs
' 実行結果: 状態・スタートアップ種別の集計処理の結果やエラーをコンソールに表示します。

' ローカルサービスの状態・スタートアップ種別の件数を集計します。構成変更はしません。
Option Explicit
Dim service, services, item, runningCount, stoppedCount, otherCount, autoCount, manualCount, disabledCount
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT State, StartMode FROM Win32_Service")
runningCount = 0
stoppedCount = 0
otherCount = 0
autoCount = 0
manualCount = 0
disabledCount = 0
For Each item In services
    Select Case LCase(CStr(item.State))
        Case "running": runningCount = runningCount + 1
        Case "stopped": stoppedCount = stoppedCount + 1
        Case Else: otherCount = otherCount + 1
    End Select
    Select Case LCase(CStr(item.StartMode))
        Case "auto": autoCount = autoCount + 1
        Case "manual": manualCount = manualCount + 1
        Case "disabled": disabledCount = disabledCount + 1
    End Select
Next
WScript.Echo "Running=" & runningCount & " Stopped=" & stoppedCount & " Other=" & otherCount
WScript.Echo "Automatic=" & autoCount & " Manual=" & manualCount & " Disabled=" & disabledCount
