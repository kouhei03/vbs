' 解説: 固定ドライブに対してデフラグ分析のみを実行し、最適化は行いません。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: fso, shell, drives, drive, process, line, found
Dim fso, shell, drives, drive, process, line, found
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso = CreateObject("Scripting.FileSystemObject")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数へ格納します。
found = False
' 解説: コレクション内の各項目を順に処理します: For Each drive In fso.Drives
For Each drive In fso.Drives
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: drive.DriveType = 2 And drive.IsReady
    If drive.DriveType = 2 And drive.IsReady Then
        ' 解説: 取得値または計算結果を変数へ格納します。
        found = True
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo "Analysis: " & drive.DriveLetter & ":"
        ' 解説: 読み取り用コマンドを実行し、結果取得用オブジェクトを設定します。
        Set process = shell.Exec("defrag.exe " & drive.DriveLetter & ": /A /V")
        ' 解説: 条件が成立している間、繰り返し処理します: Do While Not process.StdOut.AtEndOfStream
        Do While Not process.StdOut.AtEndOfStream
            ' 解説: 取得値または計算結果を変数へ格納します。
            line = process.StdOut.ReadLine
            ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
            WScript.Echo line
        ' 解説: Doループの現在の反復を終了します。
        Loop
        ' 解説: 条件が成立している間、繰り返し処理します: Do While Not process.StdErr.AtEndOfStream
        Do While Not process.StdErr.AtEndOfStream
            ' 解説: 取得値または計算結果を変数へ格納します。
            line = process.StdErr.ReadLine
            ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
            WScript.Echo "ERROR: " & line
        ' 解説: Doループの現在の反復を終了します。
        Loop
        ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: process.ExitCode <> 0
        If process.ExitCode <> 0 Then WScript.Echo "defrag analysis exit code=" & process.ExitCode
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Not found
If Not found Then WScript.Echo "No ready fixed drives found."
