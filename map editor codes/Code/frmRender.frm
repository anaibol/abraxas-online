VERSION 5.00
Begin VB.Form frmRender 
   Caption         =   "Form1"
   ClientHeight    =   12120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   808
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   810
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   12000
      Left            =   60
      ScaleHeight     =   800
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   60
      Width           =   12000
   End
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Sub Render(ByVal x As Integer, ByVal y As Integer)
    Dim miRect As RECT
    
    With miRect
        .bottom = 1600
        .Right = 1600
    End With
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
        Engine.RenderScreen
    D3DDevice.EndScene
    D3DDevice.Present miRect, miRect, picRender.hwnd, ByVal 0
                
    Engine.SaveBackBuffer App.Path & "\renders\" & y & "x" & x & ".bmp", 1600, 1600
End Sub
Public Sub Draw_Grh(ByVal desthDC As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef file_path As String)
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim PrevObj As Long
    Dim grh_index As Integer
    
    src_x = 0
    src_y = 0
    src_width = 1600
    src_height = 1600
            
    hdcsrc = CreateCompatibleDC(desthDC)
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
    BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy

    Call DeleteObject(SelectObject(hdcsrc, PrevObj))
    DeleteDC hdcsrc
End Sub

Private Sub Form_Load()
    Dim y As Long, x As Long
    
    Engine.Set_Map 24, 24
    
    picRender.Width = 1600
    picRender.Height = 1600
  
    UserPos.x = 25
    UserPos.y = 25
    Render 1, 1
    
    UserPos.x = 75
    UserPos.y = 25
    Render 2, 1
    
    UserPos.x = 75
    UserPos.y = 75
    Render 2, 2
    
    UserPos.x = 25
    UserPos.y = 75
    Render 1, 2
    
    picRender.Cls
    picRender.Width = 3200
    picRender.Height = 3200
    For y = 1 To 2
        For x = 1 To 2
            Draw_Grh picRender.hdc, (x - 1) * 1600, (y - 1) * 1600, App.Path & "\renders\" & y & "x" & x & ".bmp"
        Next x
    Next y
    
    picRender.Refresh
    
    SavePicture picRender.Image, App.Path & "\map.bmp"
    
    Engine.Engine_Reset

    MsgBox "Mapa renderizado correctamente y guardado en " & App.Path & "\map.bmp"
End Sub

