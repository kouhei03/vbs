' ファイル概要: ローカルセキュリティの集計をHTMLに出力する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo CreateSecurityAuditSummary.vbs <new-html-path>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ローカルアカウント・プロファイル・共有の集計を新規HTMLファイルに保存し、処理結果を表示します。

Option Explicit

Dim service, users, profiles, shares, user, profile, share
Dim userCount, disabledCount, lockedCount, profileCount, shareCount
Dim fileSystem, outputPath, reportHtml, reportFile

' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo CreateSecurityAuditSummary.vbs <new-html-path>"
    WScript.Quit 2
End If
outputPath = WScript.Arguments(0)
Set fileSystem = CreateObject("Scripting.FileSystemObject")
If fileSystem.FileExists(outputPath) Then
    WScript.Echo "Output file already exists; refusing to overwrite: " & outputPath
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_UserAccount を照会し、Disabled, Lockout を取得します。
Set users = service.ExecQuery( _
    "SELECT Disabled, Lockout FROM Win32_UserAccount WHERE LocalAccount=True")
' [処理解説] WMIの Win32_UserProfile を照会し、SID を取得します。
Set profiles = service.ExecQuery( _
    "SELECT SID FROM Win32_UserProfile WHERE Special=False")
' [処理解説] WMIの Win32_Share を照会し、Name を取得します。
Set shares = service.ExecQuery("SELECT Name FROM Win32_Share")

userCount = 0
disabledCount = 0
lockedCount = 0
For Each user In users
    userCount = userCount + 1
    If user.Disabled Then disabledCount = disabledCount + 1
    If user.Lockout Then lockedCount = lockedCount + 1
Next
profileCount = 0
For Each profile In profiles
    profileCount = profileCount + 1
Next
shareCount = 0
For Each share In shares
    shareCount = shareCount + 1
Next

reportHtml = "<!doctype html>" & vbCrLf & _
    "<html><head><meta charset=""utf-8""><title>Local Security Summary</title></head><body>" & vbCrLf & _
    "<h1>Local Security Inventory Summary</h1>" & vbCrLf & _
    "<p>Generated: " & ServerTime() & "</p>" & vbCrLf & _
    "<ul><li>Local user accounts: " & userCount & "</li>" & _
    "<li>Disabled accounts: " & disabledCount & "</li>" & _
    "<li>Locked accounts: " & lockedCount & "</li>" & _
    "<li>Non-special user profiles: " & profileCount & "</li>" & _
    "<li>Local shares: " & shareCount & "</li></ul>" & vbCrLf & _
    "<p>Summary counts only; no names, paths, credentials, or ACL entries are included.</p>" & _
    "</body></html>"

On Error Resume Next
' [処理解説] 上書き引数をFalseにして、新しい出力ファイルのみを作成します。
Set reportFile = fileSystem.CreateTextFile(outputPath, False, True)
If Err.Number <> 0 Then
    WScript.Echo "Could not create report: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
reportFile.Write reportHtml
reportFile.Close
WScript.Echo "Summary report written to: " & outputPath

Function ServerTime()
    ServerTime = CStr(Now)
End Function
