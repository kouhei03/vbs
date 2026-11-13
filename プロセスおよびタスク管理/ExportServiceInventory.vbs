' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、services、item、fileSystem、outputPath、outputFile）。主な用途: WMIサービス接続、サービス一覧、現在処理中のWMI項目、ファイルシステム操作オブジェクト、出力先パス、出力ファイル。
Dim service, services, item, fileSystem, outputPath, outputFile

' 解説: 条件「WScript.Arguments.Count <> 1 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo ExportServiceInventory.vbs <new-csv-path>"
    WScript.Echo "Usage: cscript //nologo ExportServiceInventory.vbs <new-csv-path>"
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
' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、Name, DisplayName, State, StartModeを取得します。
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State, StartMode FROM Win32_Service")

' 解説: 直後のCOM/ファイル操作エラーをErrで確認できるよう、一時的に次の行へ処理を継続します。
On Error Resume Next
' 解説: 指定パスに新しいローカルファイルを作成します。既存ファイルを上書きしない引数で呼び出しています。
Set outputFile = fileSystem.CreateTextFile(outputPath, False, True)
' 解説: 条件「Err.Number <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Err.Number <> 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Could not create CSV file: " & Err.Description
    WScript.Echo "Could not create CSV file: " & Err.Description
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 一時的なエラー抑制を解除し、通常のエラー処理に戻します。
On Error GoTo 0

' 解説: 現在の行を出力ファイルへ書き込みます: outputFile.WriteLine """Name"",""DisplayName"",""State"",""StartMode"""
outputFile.WriteLine """Name"",""DisplayName"",""State"",""StartMode"""
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each item In services
For Each item In services
    ' 解説: 現在の行を出力ファイルへ書き込みます: outputFile.WriteLine Csv(item.Name) & "," & Csv(item.DisplayName) & "," & Csv(item.State) & "," & Csv(item.StartMode)
    outputFile.WriteLine Csv(item.Name) & "," & Csv(item.DisplayName) & "," & _
        Csv(item.State) & "," & Csv(item.StartMode)
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: ファイルを閉じ、書き込み内容を確定してリソースを解放します。
outputFile.Close
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Service inventory written to: " & outputPath
WScript.Echo "Service inventory written to: " & outputPath

' 解説: 再利用可能な関数を定義します。関数名と引数: Csv(value)
Function Csv(value)
    ' 解説: 条件「IsNull(value) Or IsEmpty(value) Then」を評価し、成立した場合に続く分岐処理を実行します。
    If IsNull(value) Or IsEmpty(value) Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: value = ""
        value = ""
    ' 解説: 条件分岐ブロックを終了します。
    End If
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: Csv = """" & Replace(CStr(value), """", """""") & """"
    Csv = """" & Replace(CStr(value), """", """""") & """"
' 解説: 関数定義を終了します。
End Function
