' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（fileSystem、shell、windowsPath、dumpPath、miniDumpFolder、folder）。主な用途: ファイルシステム操作オブジェクト、Windows Script Hostシェル。
Dim fileSystem, shell, windowsPath, dumpPath, miniDumpFolder, folder
' 解説: このスクリプトで使う変数を宣言します（fileItem、found）。主な用途: 対象が見つかったかを示すフラグ。
Dim fileItem, found
' 解説: Scripting.FileSystemObject（ローカルファイルの存在確認や新規作成に使うオブジェクト）を生成し、変数に格納します。
Set fileSystem = CreateObject("Scripting.FileSystemObject")
' 解説: WScript.Shell（Windows環境変数や固定のローカルコマンドを扱うオブジェクト）を生成し、変数に格納します。
Set shell = CreateObject("WScript.Shell")
' 解説: SystemRootなどの環境変数参照を実際のローカルパスへ展開します。
windowsPath = shell.ExpandEnvironmentStrings("%SystemRoot%")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: dumpPath = fileSystem.BuildPath(windowsPath, "MEMORY.DMP")
dumpPath = fileSystem.BuildPath(windowsPath, "MEMORY.DMP")
' 解説: 条件「fileSystem.FileExists(dumpPath) Then」を評価し、成立した場合に続く分岐処理を実行します。
If fileSystem.FileExists(dumpPath) Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Memory dump exists: " & dumpPath
    WScript.Echo "Memory dump exists: " & dumpPath
' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
Else
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "No MEMORY.DMP found at the standard path."
    WScript.Echo "No MEMORY.DMP found at the standard path."
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: miniDumpFolder = fileSystem.BuildPath(windowsPath, "Minidump")
miniDumpFolder = fileSystem.BuildPath(windowsPath, "Minidump")
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
found = False
' 解説: 条件「fileSystem.FolderExists(miniDumpFolder) Then」を評価し、成立した場合に続く分岐処理を実行します。
If fileSystem.FolderExists(miniDumpFolder) Then
    ' 解説: 指定したローカルフォルダーの内容を列挙するため、フォルダーオブジェクトを取得します。
    Set folder = fileSystem.GetFolder(miniDumpFolder)
    ' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each fileItem In folder.Files
    For Each fileItem In folder.Files
        ' 解説: 条件「LCase(fileSystem.GetExtensionName(fileItem.Name)) = "dmp" Then」を評価し、成立した場合に続く分岐処理を実行します。
        If LCase(fileSystem.GetExtensionName(fileItem.Name)) = "dmp" Then
            ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
            found = True
            ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Minidump exists: " & fileItem.Path
            WScript.Echo "Minidump exists: " & fileItem.Path
        ' 解説: 条件分岐ブロックを終了します。
        End If
    ' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
    Next
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 条件「Not found Then WScript.Echo "No minidump files found in the standard folder."」を評価し、成立した場合に続く分岐処理を実行します。
If Not found Then WScript.Echo "No minidump files found in the standard folder."
