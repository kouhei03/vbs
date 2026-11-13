' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（days、timeFilter、service、events、eventItem、count7031、count7034）。主な用途: 対象日数、WMI日時フィルター、WMIサービス接続、イベント一覧、イベントログ項目。
Dim days, timeFilter, service, events, eventItem, count7031, count7034
' 解説: 条件「WScript.Arguments.Count <> 1 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo GetServiceCrashCounts.vbs <days 1-30>"
    WScript.Echo "Usage: cscript //nologo GetServiceCrashCounts.vbs <days 1-30>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 条件「Not ParseDays(WScript.Arguments(0), days) Then」を評価し、成立した場合に続く分岐処理を実行します。
If Not ParseDays(WScript.Arguments(0), days) Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Days must be a whole number from 1 to 30."
    WScript.Echo "Days must be a whole number from 1 to 30."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: WbemScripting.SWbemDateTime（WMI日時の変換/比較に使うオブジェクト）を生成し、変数に格納します。
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
' 解説: ローカル日時をWMIが扱うCIM日時形式に変換して日時オブジェクトへ設定します。
timeFilter.SetVarDate DateAdd("d", -days, Now), True
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsイベントログ（Win32_NTLogEvent）を読み取り専用で照会し、EventCodeを取得します。 絞り込み条件には、検証済みの実行時引数や計算した日時が挿入されます（文字列連結位置を含むため、ここでは空値として表示しません）。
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='System' " & _
    "AND SourceName='Service Control Manager' " & _
    "AND (EventCode=7031 OR EventCode=7034) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: count7031 = 0
count7031 = 0
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: count7034 = 0
count7034 = 0
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each eventItem In events
For Each eventItem In events
    ' 解説: 条件「eventItem.EventCode = 7031 Then」を評価し、成立した場合に続く分岐処理を実行します。
    If eventItem.EventCode = 7031 Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: count7031 = count7031 + 1
        count7031 = count7031 + 1
    ' 解説: 先行する条件が不成立の場合に評価する追加条件です: eventItem.EventCode = 7034 Then
    ElseIf eventItem.EventCode = 7034 Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: count7034 = count7034 + 1
        count7034 = count7034 + 1
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Unexpected service termination events (last " & days & " days)"
WScript.Echo "Unexpected service termination events (last " & days & " days)"
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Event 7031: " & count7031
WScript.Echo "Event 7031: " & count7031
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Event 7034: " & count7034
WScript.Echo "Event 7034: " & count7034
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Counts only; event message text is omitted."
WScript.Echo "Counts only; event message text is omitted."

' 解説: 再利用可能な関数を定義します。関数名と引数: ParseDays(text, value)
Function ParseDays(text, value)
    ' 解説: このスクリプトで使う変数を宣言します（i、character）。
    Dim i, character
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: ParseDays = False
    ParseDays = False
    ' 解説: 条件「Len(text) = 0 Or Len(text) > 2 Then Exit Function」を評価し、成立した場合に続く分岐処理を実行します。
    If Len(text) = 0 Or Len(text) > 2 Then Exit Function
    ' 解説: 指定範囲の繰り返し処理を開始します: For i = 1 To Len(text)
    For i = 1 To Len(text)
        ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: character = Mid(text, i, 1)
        character = Mid(text, i, 1)
        ' 解説: 条件「InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function」を評価し、成立した場合に続く分岐処理を実行します。
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    ' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
    Next
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: value = CLng(text)
    value = CLng(text)
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: ParseDays = (value >= 1 And value <= 30)
    ParseDays = (value >= 1 And value <= 30)
' 解説: 関数定義を終了します。
End Function
