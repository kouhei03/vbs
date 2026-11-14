' 外部許可リストとローカルサービス名を比較し、不一致を表示します。削除や無効化はしません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: fso, inputFile, allowed, service, services, item, path, lineText, found, allowedName
Dim fso, inputFile, allowed, service, services, item, path, lineText, found, allowedName
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 169_ServiceAllowlistAudit.vbs <allowlist-file>"
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: 入力値、取得値、または計算結果を変数に格納します。
path = WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso = CreateObject("Scripting.FileSystemObject")
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Not fso.FileExists(path) Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Allowlist file not found: " & path
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set allowed = CreateObject("Scripting.Dictionary")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
allowed.CompareMode = 1
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set inputFile = fso.OpenTextFile(path, 1, False)
' 解説: 条件に従って処理を繰り返します。
Do Until inputFile.AtEndOfStream
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    lineText = Trim(inputFile.ReadLine)
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If Len(lineText) > 0 Then allowed(lineText) = True
' 解説: ループ本体を終え、次の反復条件を確認します。
Loop
' 解説: ローカルの読み取り専用診断処理を進めます。
inputFile.Close
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State FROM Win32_Service")
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each item In services
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If Not allowed.Exists(item.Name) Then WScript.Echo "Not in allowlist: " & item.Name & vbTab & item.DisplayName & vbTab & item.State
' 解説: 次の項目または反復へ進みます。
Next
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Audit completed; no service settings were changed."
