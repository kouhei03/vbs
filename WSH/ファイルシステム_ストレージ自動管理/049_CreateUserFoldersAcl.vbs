' 概要: ユーザー別フォルダー作成とModify権限の割り当てをプレビューします。
' 実行方法: Usage: cscript //nologo 049_CreateUserFoldersAcl.vbs <root-folder> <accounts.txt> [APPLY]
' 処理結果: ユーザーフォルダーとModify ACLの作成。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
' 補足: APPLYで実行。ユーザー一覧は1行1アカウント、例: DOMAIN\\user
Option Explicit
Dim fso, root, listPath, mode, listFile, account, shell, folderPath, command, process, re
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
    WScript.Echo "Usage: cscript //nologo 049_CreateUserFoldersAcl.vbs <root-folder> <accounts.txt> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
listPath = fso.GetAbsolutePathName(WScript.Arguments(1))
mode = ""
If WScript.Arguments.Count = 3 Then mode = UCase(WScript.Arguments(2))
If Not fso.FileExists(listPath) Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.@\\-]+$"
Set shell = CreateObject("WScript.Shell")
Set listFile = fso.OpenTextFile(listPath, 1, False, -2)
Do Until listFile.AtEndOfStream
    account = Trim(listFile.ReadLine)
    If account <> "" Then
        If Not re.Test(account) Then
            listFile.Close
            WScript.Echo "Unsupported account name: " & account
            WScript.Quit 2
        End If
        folderPath = fso.BuildPath(root, Replace(account, "\", "_"))
        WScript.Echo folderPath & " <- " & account
        If mode = "APPLY" Then
            If Not fso.FolderExists(root) Then fso.CreateFolder root
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
            command = "icacls.exe """ & folderPath & """ /grant:r """ & account & ":(OI)(CI)M"""
            ' 主処理: 外部コマンドを実行し、終了コードまたは出力を確認します。
            Set process = shell.Exec(command)
            Do While Not process.StdOut.AtEndOfStream
                WScript.Echo process.StdOut.ReadLine
            Loop
            If process.ExitCode <> 0 Then
                listFile.Close
                WScript.Echo process.StdErr.ReadAll
                WScript.Quit process.ExitCode
            End If
        End If
    End If
Loop
listFile.Close
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to create folders and change ACLs."
