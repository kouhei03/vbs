' 概要: 指定先にWindowsショートカット(.lnk)を作成します。既存ファイルは上書きしません。
' 実行方法: Usage: cscript //nologo 020_CreateShortcut.vbs <target-path> <new-shortcut.lnk>
' 処理結果: `.lnk`ショートカット作成。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, target, shortcutPath, shell, shortcut
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 020_CreateShortcut.vbs <target-path> <new-shortcut.lnk>"
    WScript.Quit 2
End If
' 主処理: `.lnk`ショートカット作成。
Set fso = CreateObject("Scripting.FileSystemObject")
target = fso.GetAbsolutePathName(WScript.Arguments(0))
shortcutPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FileExists(target) And Not fso.FolderExists(target) Then WScript.Quit 2
If LCase(fso.GetExtensionName(shortcutPath)) <> "lnk" Or fso.FileExists(shortcutPath) Then
    WScript.Echo "Provide a new .lnk path."
    WScript.Quit 2
End If
If Not fso.FolderExists(fso.GetParentFolderName(shortcutPath)) Then
    WScript.Echo "Shortcut parent folder does not exist."
    WScript.Quit 2
End If
Set shell = CreateObject("WScript.Shell")
Set shortcut = shell.CreateShortcut(shortcutPath)
shortcut.TargetPath = target
shortcut.WorkingDirectory = fso.GetParentFolderName(target)
shortcut.Save
WScript.Echo "Created: " & shortcutPath
