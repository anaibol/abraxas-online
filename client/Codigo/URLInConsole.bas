Attribute VB_Name = "modURlInConsole"
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const EM_SETEVENTMASK = &H445
Const EN_LINK = &H70B
Const ENM_LINK = &H4000000
Const EM_AUTOURLDETECT = &H45B
Const EM_GETEVENTMASK = &H43B
Const GWL_WNDPROC = -4
Const WM_NOTIFY = &H4E
Const WM_LBUTTONDOWN = &H201
Const EM_GETTEXTRANGE = &H44B
 
Dim lOldProc As Long
Dim hWndRTB As Long
Dim hWndParent As Long

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type
 
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type
 
Public Sub EnableURLDetect(ByVal hWndTextbox As Long, ByVal hWndOwner As Long)
  'Don't want to subclass twice!
  If lOldProc = 0 Then
    'Subclass!
    lOldProc = SetWindowLong(hWndOwner, GWL_WNDPROC, AddressOf WndProc)
    Call SendMessage(hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0))
    Call SendMessage(hWndTextbox, EM_AUTOURLDETECT, 1, ByVal 0)
    hWndParent = hWndOwner
    hWndRTB = hWndTextbox
  End If
End Sub
 
Public Sub DisableURLDetect()
  If lOldProc Then
    Call SendMessage(hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0)
    'Reset the window procedure (stop the subclassing)
    SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
    'Set this to 0 so we can subclass again in future
    lOldProc = 0
  End If
End Sub
 
Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim uHead As NMHDR
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim sText As String
    Dim lLen As Long
    
    'Which message?
    Select Case uMsg
        Case WM_NOTIFY
            'Copy the notification header into our structure from the pointer
            CopyMemory uHead, ByVal lParam, Len(uHead)
           
            If uHead.hWndFrom = hWndRTB And uHead.code = EN_LINK Then
                CopyMemory eLink, ByVal lParam, Len(eLink)
               
                'What kind of message?
                Select Case eLink.msg
               
                    Case WM_LBUTTONDOWN
                        eText.chrg.cpMin = eLink.chrg.cpMin
                        eText.chrg.cpMax = eLink.chrg.cpMax
                        eText.lpstrText = Space$(1024)
                        lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)
                        sText = Left$(eText.lpstrText, lLen)
                        
                        If UCase$(Left(sText, Len("abraxas-online.com"))) = UCase$("abraxas-online.com") Then
                            ShellExecute hWndParent, vbNullString, "http://abraxas-online.com", vbNullString, vbNullString, SW_SHOW
                        End If
                End Select
               
            End If
           
    End Select
    WndProc = CallWindowProc(lOldProc, hWnd, uMsg, wParam, lParam)
End Function
 
