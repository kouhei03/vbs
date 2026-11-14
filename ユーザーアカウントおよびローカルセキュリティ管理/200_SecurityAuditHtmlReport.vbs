' ファイル概要: ローカルセキュリティの集計結果を新規HTMLファイルに出力する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 200_SecurityAuditHtmlReport.vbs <new-local-html-path>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ユーザー・グループ・共有・プロファイルの集計を新規HTMLファイルに保存し、作成結果を表示します。

Option Explicit

Dim outputPath, fso, service, users, groups, shares, profiles
Dim item, userCount, disabledCount, groupCount, shareCount, profileCount
Dim html, report
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 200_SecurityAuditHtmlReport.vbs <new-local-html-path>"
    WScript.Quit 2
End If
outputPath = WScript.Arguments(0)
If Left(outputPath, 2) = "\\" Then
    WScript.Echo "UNC output paths are not allowed; choose a local destination."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(outputPath) Or fso.FolderExists(outputPath) Then
    WScript.Echo "Output path already exists; refusing to overwrite."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set users = service.ExecQuery( _
    "SELECT Disabled FROM Win32_UserAccount WHERE LocalAccount=True")
Set groups = service.ExecQuery( _
    "SELECT Name FROM Win32_Group WHERE LocalAccount=True")
Set shares = service.ExecQuery("SELECT Name FROM Win32_Share")
Set profiles = service.ExecQuery( _
    "SELECT SID FROM Win32_UserProfile WHERE Special=False")
userCount = 0
disabledCount = 0
For Each item In users
    userCount = userCount + 1
    If item.Disabled Then disabledCount = disabledCount + 1
Next
groupCount = 0
For Each item In groups
    groupCount = groupCount + 1
Next
shareCount = 0
For Each item In shares
    shareCount = shareCount + 1
Next
profileCount = 0
For Each item In profiles
    profileCount = profileCount + 1
Next

html = "<!doctype html><html><head><meta charset=""utf-8"">" & _
    "<title>Local Security Audit</title></head><body>" & _
    "<h1>Local Security Audit Summary</h1><p>Generated: " & _
    ServerTime() & "</p><ul><li>Local users: " & userCount & _
    "</li><li>Disabled users: " & disabledCount & _
    "</li><li>Local groups: " & groupCount & _
    "</li><li>Shares: " & shareCount & _
    "</li><li>Non-special profiles: " & profileCount & _
    "</li></ul><p>Counts only; no names, paths, credentials, or ACLs included.</p></body></html>"
On Error Resume Next
Set report = fso.CreateTextFile(outputPath, False, True)
If Err.Number <> 0 Then
    WScript.Echo "Could not create report: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
report.Write html
report.Close
WScript.Echo "Local summary report created: " & outputPath

Function ServerTime()
    ServerTime = CStr(Now)
End Function
