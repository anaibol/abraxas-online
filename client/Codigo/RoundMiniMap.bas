Attribute VB_Name = "modRoundMiniMap"
Option Explicit
 
Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal X3 As Long, _
    ByVal Y3 As Long) As Long
 
Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long
 
Public Sub Round_Picture(El_Form As PictureBox, Radio As Long)

    Dim Region As Long
    Dim Ret As Long
    Dim Ancho As Long
    Dim Alto As Long
    Dim old_Scale As Integer
   
    old_Scale = El_Form.ScaleMode
    El_Form.ScaleMode = vbPixels
    Ancho = El_Form.ScaleWidth
    Alto = El_Form.ScaleHeight
    Region = CreateRoundRectRgn(0, 0, Ancho, Alto, Radio, Radio)
    Ret = SetWindowRgn(El_Form.hWnd, Region, True)
    El_Form.ScaleMode = old_Scale
End Sub


