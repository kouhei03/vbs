' 概要: ローカル印刷キューに存在するジョブを表示します。キュー削除はしません。
' 実行方法: cscript //nologo 027_PrintQueueAudit.vbs
' 処理結果: 印刷キュー照会（ジョブ削除なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, jobs, job
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set jobs = service.ExecQuery("SELECT Name, JobId, Owner, Document, Status FROM Win32_PrintJob")
For Each job In jobs
    WScript.Echo "PrinterJob=" & job.Name & " JobId=" & job.JobId & " Owner=" & job.Owner & _
        " Document=" & job.Document & " Status=" & job.Status
Next
If jobs.Count = 0 Then WScript.Echo "No print jobs reported."
