' ファイル概要: イベントのメタデータをCSVに出力する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ExportEventSummary.vbs <System|Application|Security> <hours 1-168> <new-csv-path>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: イベントの時刻・ID・ソース・レコード番号を新規CSVに保存し、処理結果を表示します。

Option Explicit

Dim logName, hours, outputPath, timeFilter, service, events, eventItem
Dim fileSystem, outputFile

' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo ExportEventSummary.vbs <System|Application|Security> <hours 1-168> <new-csv-path>"
    WScript.Quit 2
End If
logName = WScript.Arguments(0)
If Not IsAllowedLog(logName) Then
    WScript.Echo "Log must be System, Application, or Security."
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(1), hours) Then
    WScript.Echo "Hours must be a whole number from 1 to 168."
    WScript.Quit 2
End If

outputPath = WScript.Arguments(2)
Set fileSystem = CreateObject("Scripting.FileSystemObject")
If fileSystem.FileExists(outputPath) Then
    WScript.Echo "Output file already exists; refusing to overwrite."
    WScript.Quit 2
End If

Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
timeFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, SourceName, TimeGenerated, RecordNumber を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='" & logName & _
    "' AND TimeGenerated >= '" & timeFilter.Value & "'")

On Error Resume Next
' [処理解説] 上書き引数をFalseにして、新しい出力ファイルのみを作成します。
Set outputFile = fileSystem.CreateTextFile(outputPath, False, True)
If Err.Number <> 0 Then
    WScript.Echo "Could not create CSV file: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

outputFile.WriteLine """Time"",""EventID"",""Source"",""RecordNumber"""
For Each eventItem In events
    outputFile.WriteLine Csv(EventTime(eventItem.TimeGenerated)) & "," & _
        Csv(eventItem.EventCode) & "," & Csv(eventItem.SourceName) & "," & _
        Csv(eventItem.RecordNumber)
Next
outputFile.Close
WScript.Echo "Event metadata exported; event message text was not included."

Function IsAllowedLog(name)
    IsAllowedLog = (StrComp(name, "System", vbTextCompare) = 0 Or _
        StrComp(name, "Application", vbTextCompare) = 0 Or _
        StrComp(name, "Security", vbTextCompare) = 0)
End Function

' [処理解説] 数字以外の入力を除外してから整数へ変換し、指定範囲に収まるか判定します。
Function ParseHours(text, value)
    Dim i, ch
    ParseHours = False
    If Len(text) = 0 Or Len(text) > 3 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If InStr(1, "0123456789", ch, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseHours = (value >= 1 And value <= 168)
End Function

Function EventTime(cimTime)
    Dim dateTime
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    dateTime.Value = cimTime
    EventTime = CStr(dateTime.GetVarDate(True))
End Function

' [処理解説] CSVの値を引用符で囲み、値中の引用符を二重化します。
Function Csv(value)
    If IsNull(value) Or IsEmpty(value) Then value = ""
    Csv = """" & Replace(CStr(value), """", """""") & """"
End Function
