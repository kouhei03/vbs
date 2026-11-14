' ローカル印刷キューに存在するジョブを表示します。キュー削除はしません。
Option Explicit
Dim service, jobs, job
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set jobs = service.ExecQuery("SELECT Name, JobId, Owner, Document, Status FROM Win32_PrintJob")
For Each job In jobs
    WScript.Echo "PrinterJob=" & job.Name & " JobId=" & job.JobId & " Owner=" & job.Owner & _
        " Document=" & job.Document & " Status=" & job.Status
Next
If jobs.Count = 0 Then WScript.Echo "No print jobs reported."
