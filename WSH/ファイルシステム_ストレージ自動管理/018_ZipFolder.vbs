' 概要: ShellのZIP機能でフォルダーを圧縮します。既存のZIPは上書きしません。
' 実行方法: Usage: cscript //nologo 018_ZipFolder.vbs <folder> <new-archive.zip>
' 処理結果: Windows Shellによる新規ZIP作成。処理結果や状態はコンソールに表示されます。
' コード解説: ADODB.Streamでデータを読み書きし、指定形式の結果を扱います。
Option Explicit
Dim fso, source, zipPath, stream, shell, zipFolder, sourceFolder, deadline, lastSize, stableSince, currentSize
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 018_ZipFolder.vbs <folder> <new-archive.zip>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
source = fso.GetAbsolutePathName(WScript.Arguments(0))
zipPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FolderExists(source) Or LCase(fso.GetExtensionName(zipPath)) <> "zip" Or _
    fso.FileExists(zipPath) Or fso.FolderExists(zipPath) Or Not fso.FolderExists(fso.GetParentFolderName(zipPath)) Then WScript.Quit 2
If IsSameOrChild(zipPath, source) Then
    WScript.Echo "Archive must be outside the source folder."
    WScript.Quit 2
End If
Set stream = CreateObject("ADODB.Stream")
stream.Type = 2
stream.Charset = "windows-1252"
stream.Open
stream.WriteText Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, Chr(0))
stream.SaveToFile zipPath, 1
stream.Close
Set shell = CreateObject("Shell.Application")
Set zipFolder = shell.NameSpace(zipPath)
Set sourceFolder = shell.NameSpace(source)
zipFolder.CopyHere sourceFolder.Items, 20
deadline = DateAdd("n", 10, Now)
lastSize = -1
stableSince = Now
' 主処理: 条件を確認しながら処理を繰り返します。
Do While Now < deadline
    currentSize = fso.GetFile(zipPath).Size
    If currentSize <> lastSize Then
        lastSize = currentSize
        stableSince = Now
    ElseIf zipFolder.Items.Count >= sourceFolder.Items.Count And DateDiff("s", stableSince, Now) >= 5 Then
        Exit Do
    End If
    WScript.Sleep 1000
Loop
If Now >= deadline Or zipFolder.Items.Count < sourceFolder.Items.Count Then
    WScript.Echo "Archive operation timed out; inspect " & zipPath
    WScript.Quit 1
End If
WScript.Echo "Created: " & zipPath
Function IsSameOrChild(path, parentPath)
    Dim normalizedPath, normalizedParent
    normalizedPath = LCase(path)
    normalizedParent = LCase(parentPath)
    If Right(normalizedPath, 1) = "\" And Len(normalizedPath) > 3 Then normalizedPath = Left(normalizedPath, Len(normalizedPath) - 1)
    If Right(normalizedParent, 1) = "\" And Len(normalizedParent) > 3 Then normalizedParent = Left(normalizedParent, Len(normalizedParent) - 1)
    If normalizedPath = normalizedParent Then
        IsSameOrChild = True
    ElseIf Right(normalizedParent, 1) = "\" Then
        IsSameOrChild = (Left(normalizedPath, Len(normalizedParent)) = normalizedParent)
    Else
        IsSameOrChild = (Left(normalizedPath, Len(normalizedParent) + 1) = normalizedParent & "\")
    End If
End Function
