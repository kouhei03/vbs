' ローカル全サービスの状態を、新規CSVファイルへ一度だけ出力します。既存ファイルは上書きしません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: fso, path, file, service, services, item
Dim fso, path, file, service, services, item
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 181_ServiceStateSnapshot.vbs <new-csv-path>"
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: 入力値、取得値、または計算結果を変数に格納します。
path = WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso = CreateObject("Scripting.FileSystemObject")
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If fso.FileExists(path) Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Refusing to overwrite: " & path
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode, StartName, ExitCode FROM Win32_Service")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set file = fso.CreateTextFile(path, False, True)
' 解説: ローカルの読み取り専用診断処理を進めます。
file.WriteLine "Name,DisplayName,State,StartMode,StartName,ExitCode"
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each item In services
    ' 解説: ローカルの読み取り専用診断処理を進めます。
    file.WriteLine Csv(item.Name) & "," & Csv(item.DisplayName) & "," & Csv(item.State) & "," & Csv(item.StartMode) & "," & Csv(item.StartName) & "," & Csv(item.ExitCode)
' 解説: 次の項目または反復へ進みます。
Next
' 解説: ローカルの読み取り専用診断処理を進めます。
file.Close
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Snapshot written: " & path
' 解説: 再利用する補助関数を定義します。
Function Csv(value)
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If IsNull(value) Or IsEmpty(value) Then value = ""
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    Csv = """" & Replace(CStr(value), """", """""") & """"
' 解説: 補助関数の定義を終了します。
End Function
