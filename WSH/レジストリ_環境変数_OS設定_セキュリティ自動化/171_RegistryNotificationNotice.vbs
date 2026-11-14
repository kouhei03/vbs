' レジストリ変更通知は送信しません。設定を変更していないことを明示する案内です。
Option Explicit
WScript.Echo "No registry change was made; no WM_SETTINGCHANGE notification was sent."
WScript.Echo "Only notify applications after an intentional, approved setting change."
