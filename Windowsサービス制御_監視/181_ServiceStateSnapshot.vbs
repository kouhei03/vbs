' ファイル概要: 新規CSVへのサービス状態スナップショット。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 181_ServiceStateSnapshot.vbs <new-csv-path>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: サービス状態のCSVファイルを新規作成し、処理結果を表示します。

' ローカル全サービスの状態を、新規CSVファイルへ一度だけ出力します。既存ファイルは上書きしません。
Option Explicit
Dim fso, path, file, service, services, item
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 181_ServiceStateSnapshot.vbs <new-csv-path>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then
    WScript.Echo "Refusing to overwrite: " & path
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State, StartMode, StartName, ExitCode を取得します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode, StartName, ExitCode FROM Win32_Service")
' [処理解説] 上書き引数をFalseにして、新しい出力ファイルのみを作成します。
Set file = fso.CreateTextFile(path, False, True)
file.WriteLine "Name,DisplayName,State,StartMode,StartName,ExitCode"
For Each item In services
    file.WriteLine Csv(item.Name) & "," & Csv(item.DisplayName) & "," & Csv(item.State) & "," & Csv(item.StartMode) & "," & Csv(item.StartName) & "," & Csv(item.ExitCode)
Next
file.Close
WScript.Echo "Snapshot written: " & path
' [処理解説] CSVの値を引用符で囲み、値中の引用符を二重化します。
Function Csv(value)
    If IsNull(value) Or IsEmpty(value) Then value = ""
    Csv = """" & Replace(CStr(value), """", """""") & """"
End Function
