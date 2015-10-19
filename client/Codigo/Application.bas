Attribute VB_Name = "modApplication"
Option Explicit

'Retrieves the active window's hWnd for this app.
Private Declare Function GetActiveWindow Lib "user32" () As Long

'Checks if this is the active (foreground) application or not.
Public Function IsAppActive() As Boolean
    IsAppActive = (GetActiveWindow > 0)
End Function
