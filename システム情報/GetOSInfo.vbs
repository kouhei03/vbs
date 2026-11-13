' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（wmiService）。
Dim wmiService
' 解説: このスクリプトで使う変数を宣言します（operatingSystems）。
Dim operatingSystems
' 解説: このスクリプトで使う変数を宣言します（operatingSystem）。
Dim operatingSystem

' 解説: 直後のCOM/ファイル操作エラーをErrで確認できるよう、一時的に次の行へ処理を継続します。
On Error Resume Next
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 条件「Err.Number <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Err.Number <> 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Failed to connect to WMI: " & Err.Description
    WScript.Echo "Failed to connect to WMI: " & Err.Description
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 前回のCOM操作エラー情報を消去し、次の操作で発生するエラーを正確に判定できるようにします。
Err.Clear
' 解説: WMIのWindows OS（Win32_OperatingSystem）を読み取り専用で照会し、Caption, Version, BuildNumberを取得します。
Set operatingSystems = wmiService.ExecQuery( _
    "SELECT Caption, Version, BuildNumber FROM Win32_OperatingSystem")
' 解説: 条件「Err.Number <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Err.Number <> 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Failed to query operating system information: " & Err.Description
    WScript.Echo "Failed to query operating system information: " & Err.Description
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 一時的なエラー抑制を解除し、通常のエラー処理に戻します。
On Error GoTo 0

' 解説: 条件「operatingSystems.Count = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If operatingSystems.Count = 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "No operating system information was returned."
    WScript.Echo "No operating system information was returned."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each operatingSystem In operatingSystems
For Each operatingSystem In operatingSystems
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "OS Name: " & operatingSystem.Caption
    WScript.Echo "OS Name: " & operatingSystem.Caption
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Version: " & operatingSystem.Version
    WScript.Echo "Version: " & operatingSystem.Version
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Build Number: " & operatingSystem.BuildNumber
    WScript.Echo "Build Number: " & operatingSystem.BuildNumber
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
