' 概要: 現在の既定プリンターを表示します。プリンター設定は変更しません。
' 実行方法: cscript //nologo 008_DefaultPrinterAudit.vbs
' 処理結果: 既定プリンター確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, printers, printer, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set printers = service.ExecQuery("SELECT Name, Default, Network FROM Win32_Printer WHERE Default=True")
found = False
For Each printer In printers
    WScript.Echo "DefaultPrinter=" & printer.Name & " Network=" & printer.Network
    found = True
Next
If Not found Then WScript.Echo "No default printer reported."
