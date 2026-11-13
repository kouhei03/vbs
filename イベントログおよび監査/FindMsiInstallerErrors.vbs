' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（timeFilter、service、events、eventItem）。主な用途: WMI日時フィルター、WMIサービス接続、イベント一覧、イベントログ項目。
Dim timeFilter, service, events, eventItem
' 解説: WbemScripting.SWbemDateTime（WMI日時の変換/比較に使うオブジェクト）を生成し、変数に格納します。
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
' 解説: ローカル日時をWMIが扱うCIM日時形式に変換して日時オブジェクトへ設定します。
timeFilter.SetVarDate DateAdd("h", -168, Now), True
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsイベントログ（Win32_NTLogEvent）を読み取り専用で照会し、EventCode, SourceName, TimeGenerated, RecordNumberを取得します。 絞り込み条件には、検証済みの実行時引数や計算した日時が挿入されます（文字列連結位置を含むため、ここでは空値として表示しません）。
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='Application' " & _
    "AND SourceName='MsiInstaller' AND Type='Error' " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "MsiInstaller errors (last 7 days)"
WScript.Echo "MsiInstaller errors (last 7 days)"
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Time" & vbTab & "Event ID" & vbTab & "Record"
WScript.Echo "Time" & vbTab & "Event ID" & vbTab & "Record"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each eventItem In events
For Each eventItem In events
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: EventTime(eventItem.TimeGenerated) & vbTab & eventItem.EventCode & vbTab & eventItem.RecordNumber
    WScript.Echo EventTime(eventItem.TimeGenerated) & vbTab & _
        eventItem.EventCode & vbTab & eventItem.RecordNumber
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 再利用可能な関数を定義します。関数名と引数: EventTime(cimTime)
Function EventTime(cimTime)
    ' 解説: このスクリプトで使う変数を宣言します（dateTime）。主な用途: WMI日時変換オブジェクト。
    Dim dateTime
    ' 解説: WbemScripting.SWbemDateTime（WMI日時の変換/比較に使うオブジェクト）を生成し、変数に格納します。
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    ' 解説: WMIから得たCIM日時文字列を日時変換オブジェクトに設定し、VBScript日付型へ変換できるようにします。
    dateTime.Value = cimTime
    ' 解説: CIM日時値をVBScriptの日付型へ変換し、表示または計算に使います。
    EventTime = CStr(dateTime.GetVarDate(True))
' 解説: 関数定義を終了します。
End Function
