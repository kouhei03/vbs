' ローカルプロセスの件数をSessionIdごとに集計します。
Option Explicit
Dim service, processes, process, counts, sessionKey
Set service = GetObject("winmgmts:\\.\root\cimv2")
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
