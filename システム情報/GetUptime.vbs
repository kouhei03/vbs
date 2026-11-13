' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、systems、system、dateTime、bootTime、elapsedMinutes）。主な用途: WMIサービス接続、OS情報一覧、OS情報、WMI日時変換オブジェクト、最終起動日時、経過時間（分）。
Dim service, systems, system, dateTime, bootTime, elapsedMinutes
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindows OS（Win32_OperatingSystem）を読み取り専用で照会し、LastBootUpTimeを取得します。
Set systems = service.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
' 解説: WbemScripting.SWbemDateTime（WMI日時の変換/比較に使うオブジェクト）を生成し、変数に格納します。
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each system In systems
For Each system In systems
    ' 解説: WMIから得たCIM日時文字列を日時変換オブジェクトに設定し、VBScript日付型へ変換できるようにします。
    dateTime.Value = system.LastBootUpTime
    ' 解説: CIM日時値をVBScriptの日付型へ変換し、表示または計算に使います。
    bootTime = dateTime.GetVarDate(True)
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: elapsedMinutes = CLng((Now - bootTime) * 1440)
    elapsedMinutes = CLng((Now - bootTime) * 1440)
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Last boot: " & CStr(bootTime)
    WScript.Echo "Last boot: " & CStr(bootTime)
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Uptime: " & (elapsedMinutes \ 1440) & " days, " & ((elapsedMinutes Mod 1440) \ 60) & " hours, " & (elapsedMinutes Mod 60) & " minutes"
    WScript.Echo "Uptime: " & (elapsedMinutes \ 1440) & " days, " & _
        ((elapsedMinutes Mod 1440) \ 60) & " hours, " & _
        (elapsedMinutes Mod 60) & " minutes"
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
