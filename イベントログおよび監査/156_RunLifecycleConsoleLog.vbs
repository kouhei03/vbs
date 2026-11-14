Option Explicit

Dim startedAt, endedAt
startedAt = Now
WScript.Echo "Script started: " & CStr(startedAt)

' Place the script's normal work between these two lifecycle messages.

endedAt = Now
WScript.Echo "Script ended: " & CStr(endedAt)
WScript.Echo "This example writes to the console only; it creates no event source or event."
