Option Explicit
Dim fso, service, processes, log, path
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 167_AppendProcessLog.vbs <log-file>": WScript.Quit 2
path=WScript.Arguments(0)
Set fso=CreateObject("Scripting.FileSystemObject")
Set service=GetObject("winmgmts:\\.\root\cimv2")
Set processes=service.ExecQuery("SELECT ProcessId FROM Win32_Process")
Set log=fso.OpenTextFile(path,8,True)
log.WriteLine Now & ",process-count=" & processes.Count
log.Close
WScript.Echo "A local summary line was appended."
