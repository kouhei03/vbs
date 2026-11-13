' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、users、profiles、shares、user、profile、share）。主な用途: WMIサービス接続、ユーザー一覧、プロファイル一覧、共有一覧、ユーザー情報、ユーザープロファイル、共有情報。
Dim service, users, profiles, shares, user, profile, share
' 解説: このスクリプトで使う変数を宣言します（userCount、disabledCount、lockedCount、profileCount、shareCount）。
Dim userCount, disabledCount, lockedCount, profileCount, shareCount
' 解説: このスクリプトで使う変数を宣言します（fileSystem、outputPath、reportHtml、reportFile）。主な用途: ファイルシステム操作オブジェクト、出力先パス。
Dim fileSystem, outputPath, reportHtml, reportFile

' 解説: 条件「WScript.Arguments.Count <> 1 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo CreateSecurityAuditSummary.vbs <new-html-path>"
    WScript.Echo "Usage: cscript //nologo CreateSecurityAuditSummary.vbs <new-html-path>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: outputPath = WScript.Arguments(0)
outputPath = WScript.Arguments(0)
' 解説: Scripting.FileSystemObject（ローカルファイルの存在確認や新規作成に使うオブジェクト）を生成し、変数に格納します。
Set fileSystem = CreateObject("Scripting.FileSystemObject")
' 解説: 条件「fileSystem.FileExists(outputPath) Then」を評価し、成立した場合に続く分岐処理を実行します。
If fileSystem.FileExists(outputPath) Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Output file already exists; refusing to overwrite: " & outputPath
    WScript.Echo "Output file already exists; refusing to overwrite: " & outputPath
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのユーザーアカウント（Win32_UserAccount）を読み取り専用で照会し、Disabled, Lockoutを取得します。 固定絞り込み条件: LocalAccount=True。
Set users = service.ExecQuery( _
    "SELECT Disabled, Lockout FROM Win32_UserAccount WHERE LocalAccount=True")
' 解説: WMIのユーザープロファイル（Win32_UserProfile）を読み取り専用で照会し、SIDを取得します。 固定絞り込み条件: Special=False。
Set profiles = service.ExecQuery( _
    "SELECT SID FROM Win32_UserProfile WHERE Special=False")
' 解説: WMIの共有フォルダー（Win32_Share）を読み取り専用で照会し、Nameを取得します。
Set shares = service.ExecQuery("SELECT Name FROM Win32_Share")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: userCount = 0
userCount = 0
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: disabledCount = 0
disabledCount = 0
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: lockedCount = 0
lockedCount = 0
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each user In users
For Each user In users
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: userCount = userCount + 1
    userCount = userCount + 1
    ' 解説: 条件「user.Disabled Then disabledCount = disabledCount + 1」を評価し、成立した場合に続く分岐処理を実行します。
    If user.Disabled Then disabledCount = disabledCount + 1
    ' 解説: 条件「user.Lockout Then lockedCount = lockedCount + 1」を評価し、成立した場合に続く分岐処理を実行します。
    If user.Lockout Then lockedCount = lockedCount + 1
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: profileCount = 0
profileCount = 0
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each profile In profiles
For Each profile In profiles
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: profileCount = profileCount + 1
    profileCount = profileCount + 1
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: shareCount = 0
shareCount = 0
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each share In shares
For Each share In shares
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: shareCount = shareCount + 1
    shareCount = shareCount + 1
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: reportHtml = "<!doctype html>" & vbCrLf & "<html><head><meta charset=""utf-8""><title>Local Security Summary</title></head><body>" & vbCrLf & "<h1>Loc
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

' 解説: 直後のCOM/ファイル操作エラーをErrで確認できるよう、一時的に次の行へ処理を継続します。
On Error Resume Next
' 解説: 指定パスに新しいローカルファイルを作成します。既存ファイルを上書きしない引数で呼び出しています。
Set reportFile = fileSystem.CreateTextFile(outputPath, False, True)
' 解説: 条件「Err.Number <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Err.Number <> 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Could not create report: " & Err.Description
    WScript.Echo "Could not create report: " & Err.Description
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 一時的なエラー抑制を解除し、通常のエラー処理に戻します。
On Error GoTo 0
' 解説: 生成済みの内容を出力ファイルへ書き込みます。
reportFile.Write reportHtml
' 解説: ファイルを閉じ、書き込み内容を確定してリソースを解放します。
reportFile.Close
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Summary report written to: " & outputPath
WScript.Echo "Summary report written to: " & outputPath

' 解説: 再利用可能な関数を定義します。関数名と引数: ServerTime()
Function ServerTime()
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: ServerTime = CStr(Now)
    ServerTime = CStr(Now)
' 解説: 関数定義を終了します。
End Function
