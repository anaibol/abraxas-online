Attribute VB_Name = "modRichTextBox"
Public Vertical_Pos As Long
Public MouseOverRecTxt As Boolean

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
 
Private Const EM_GETTHUMB = &HBE
Private Const SB_THUMBPOSITION = &H4
Private Const WM_VSCROLL = &H115

Public Function GetVerticalScrollPos(rtb As RichTextBox) As Long
  GetVerticalScrollPos = SendMessage(rtb.hWnd, EM_GETTHUMB, 0&, 0&)
End Function

Public Sub SetVerticalScrollPos(rtb As RichTextBox, Position As Long)
  SendMessage rtb.hWnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * Position, Nothing
End Sub

Public Function GetScrollBarInfo(hWnd As Long) As SCROLLINFO
    Dim SBI As SCROLLINFO
    SBI.cbSize = Len(SBI)
    SBI.fMask = SIF_ALL
    GetScrollInfo hWnd, SB_HORZ, SBI
    GetScrollBarInfo = SBI
End Function

Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal Red As Byte = 255, Optional ByVal Green As Byte = 255, Optional ByVal Blue As Byte = 255, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False, Optional ByVal IsDialog As Boolean = False)
    
    On Error Resume Next
    
    'Call DibujarConsola(Text, Red, Green, Blue, bold, italic, 0)
    
    Vertical_Pos = GetVerticalScrollPos(frmMain.RecTxt)
    
    With frmMain.RecTxt
    
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not Red = -1 Then
            .SelColor = RGB(Red, Green, Blue)
        End If
        
        If IsDialog Then
            .SelFontSize = 8
            '.SelItalic = False
        Else
            .SelFontSize = 9
            '.SelItalic = True
        End If
        
        .SelText = IIf(bCrLf, Message, Message & vbCrLf)
    End With
End Sub

Public Sub RenderCompas()
    
    On Error Resume Next
    
    Vertical_Pos = GetVerticalScrollPos(frmMain.CompaRecTxt)
    
    With frmMain.CompaRecTxt
        .Text = vbNullString
        
        .SelStart = 0
        .SelLength = 0
        
        Dim i As Byte
                
        For i = 1 To MaxCompaSlots
            If Compa(i).Online Then
                If LenB(Compa(i).Nombre) > 0 Then
                    .SelColor = RGB(0, 180, 0)
                    .SelText = Compa(i).Nombre & vbCrLf
                End If
            End If
        Next i
        
        For i = 1 To MaxCompaSlots
            If Not Compa(i).Online Then
                If LenB(Compa(i).Nombre) > 0 Then
                    .SelColor = RGB(180, 0, 0)
                    .SelText = Compa(i).Nombre & vbCrLf
                End If
            End If
        Next i
        
    End With
End Sub
