Option Explicit

Dim network
Set network = CreateObject("WScript.Network")
WScript.Echo "Computer name: " & network.ComputerName
