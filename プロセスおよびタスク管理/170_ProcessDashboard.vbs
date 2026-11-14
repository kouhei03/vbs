Option Explicit
Dim fso, service, processes, process, report, path
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 170_ProcessDashboard.vbs <new-html-path>": WScript.Quit 2
path=WScript.Arguments(0)
Set fso=CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then WScript.Echo "Refusing to overwrite existing file.": WScript.Quit 1
Set service=GetObject("winmgmts:\\.\root\cimv2")
Set processes=service.ExecQuery("SELECT Name,ProcessId FROM Win32_Process")
Set report=fso.CreateTextFile(path,False,True)
report.WriteLine "<!doctype html><html><head><meta charset=""utf-8""><title>Process inventory</title></head><body>"
report.WriteLine "<h1>Local process inventory</h1><p>Generated " & Html(Now) & "</p><table><tr><th>Name</th><th>PID</th></tr>"
For Each process In processes
    report.WriteLine "<tr><td>" & Html(process.Name) & "</td><td>" & process.ProcessId & "</td></tr>"
Next
report.WriteLine "</table></body></html>"
report.Close
WScript.Echo "Dashboard written: " & path

Function Html(value)
    Dim text
    text=CStr(value)
    text=Replace(text,"&","&amp;")
    text=Replace(text,"<","&lt;")
    text=Replace(text,">","&gt;")
    text=Replace(text,Chr(34),"&quot;")
    Html=text
End Function
