' 概要: ローカルプロセスの件数をSessionIdごとに集計します。
' 実行方法: cscript //nologo 083_ProcessUseBySession.vbs
' 処理結果: SessionId別プロセス数の集計。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, processes, process, counts, sessionKey
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set processes = service.ExecQuery("SELECT SessionId FROM Win32_Process")
Set counts = CreateObject("Scripting.Dictionary")
For Each process In processes
    sessionKey = CStr(process.SessionId)
    If Not counts.Exists(sessionKey) Then counts.Add sessionKey, 0
    counts(sessionKey) = counts(sessionKey) + 1
Next
For Each sessionKey In counts.Keys
    WScript.Echo "SessionId=" & sessionKey & " ProcessCount=" & counts(sessionKey)
Next
