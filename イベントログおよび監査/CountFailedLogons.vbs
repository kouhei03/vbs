' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（hours、timeFilter、service、events、eventItem、count）。主な用途: 対象時間数、WMI日時フィルター、WMIサービス接続、イベント一覧、イベントログ項目、一致件数。
Dim hours, timeFilter, service, events, eventItem, count

' 解説: 条件「WScript.Arguments.Count <> 1 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo CountFailedLogons.vbs <hours 1-168>"
    WScript.Echo "Usage: cscript //nologo CountFailedLogons.vbs <hours 1-168>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 条件「Not ParseHours(WScript.Arguments(0), hours) Then」を評価し、成立した場合に続く分岐処理を実行します。
If Not ParseHours(WScript.Arguments(0), hours) Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Hours must be a whole number from 1 to 168."
    WScript.Echo "Hours must be a whole number from 1 to 168."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: WbemScripting.SWbemDateTime（WMI日時の変換/比較に使うオブジェクト）を生成し、変数に格納します。
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
' 解説: ローカル日時をWMIが扱うCIM日時形式に変換して日時オブジェクトへ設定します。
timeFilter.SetVarDate DateAdd("h", -hours, Now), True
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsイベントログ（Win32_NTLogEvent）を読み取り専用で照会し、EventCodeを取得します。 絞り込み条件には、検証済みの実行時引数や計算した日時が挿入されます（文字列連結位置を含むため、ここでは空値として表示しません）。
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='Security' " & _
    "AND EventCode=4625 AND TimeGenerated >= '" & timeFilter.Value & "'")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: count = 0
count = 0
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each eventItem In events
For Each eventItem In events
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: count = count + 1
    count = count + 1
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Failed logon events (4625) in the last " & hours & " hours: " & count
WScript.Echo "Failed logon events (4625) in the last " & hours & _
    " hours: " & count
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Only a count is displayed; account and network details are omitted."
WScript.Echo "Only a count is displayed; account and network details are omitted."

' 解説: 再利用可能な関数を定義します。関数名と引数: ParseHours(text, value)
Function ParseHours(text, value)
    ' 解説: このスクリプトで使う変数を宣言します（i、ch）。
    Dim i, ch
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: ParseHours = False
    ParseHours = False
    ' 解説: 条件「Len(text) = 0 Or Len(text) > 3 Then Exit Function」を評価し、成立した場合に続く分岐処理を実行します。
    If Len(text) = 0 Or Len(text) > 3 Then Exit Function
    ' 解説: 指定範囲の繰り返し処理を開始します: For i = 1 To Len(text)
    For i = 1 To Len(text)
        ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: ch = Mid(text, i, 1)
        ch = Mid(text, i, 1)
        ' 解説: 条件「InStr(1, "0123456789", ch, vbBinaryCompare) = 0 Then Exit Function」を評価し、成立した場合に続く分岐処理を実行します。
        If InStr(1, "0123456789", ch, vbBinaryCompare) = 0 Then Exit Function
    ' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
    Next
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: value = CLng(text)
    value = CLng(text)
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: ParseHours = (value >= 1 And value <= 168)
    ParseHours = (value >= 1 And value <= 168)
' 解説: 関数定義を終了します。
End Function
