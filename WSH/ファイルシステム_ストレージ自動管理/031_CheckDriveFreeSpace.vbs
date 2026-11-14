' 概要: 指定ドライブの空き容量を確認し、閾値を下回ると警告します。
' 実行方法: Usage: cscript //nologo 031_CheckDriveFreeSpace.vbs <drive-letter> <minimum-free-MB>
' 処理結果: ドライブ空き容量と閾値の確認。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, driveName, threshold, drive
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 031_CheckDriveFreeSpace.vbs <drive-letter> <minimum-free-MB>"
    WScript.Quit 2
End If
' 主処理: ドライブ空き容量と閾値の確認。
If Not IsNumeric(WScript.Arguments(1)) Then
    WScript.Echo "Minimum free-space threshold must be numeric."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
driveName = UCase(Replace(WScript.Arguments(0), ":", "")) & ":"
If Len(driveName) <> 2 Or CDbl(WScript.Arguments(1)) < 0 Then WScript.Quit 2
threshold = CDbl(WScript.Arguments(1)) * 1048576
If Not fso.DriveExists(driveName) Then
    WScript.Echo "Drive not found: " & driveName
    WScript.Quit 1
End If
Set drive = fso.GetDrive(driveName)
WScript.Echo "Drive=" & driveName & " FreeBytes=" & drive.FreeSpace
If drive.FreeSpace < threshold Then
    WScript.Echo "WARNING: free space is below " & WScript.Arguments(1) & " MB."
    WScript.Quit 1
End If
