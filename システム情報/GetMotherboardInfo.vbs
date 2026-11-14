Option Explicit

Dim service, boards, board
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set boards = service.ExecQuery( _
    "SELECT Manufacturer, Product FROM Win32_BaseBoard")

For Each board In boards
    WScript.Echo "Manufacturer: " & board.Manufacturer
    WScript.Echo "Model: " & board.Product
Next
