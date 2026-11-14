' 現在の既定プリンターを表示します。プリンター設定は変更しません。
Option Explicit
Dim service, printers, printer, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set printers = service.ExecQuery("SELECT Name, Default, Network FROM Win32_Printer WHERE Default=True")
found = False
For Each printer In printers
    WScript.Echo "DefaultPrinter=" & printer.Name & " Network=" & printer.Network
    found = True
Next
If Not found Then WScript.Echo "No default printer reported."
