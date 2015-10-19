VERSION 5.00
Begin VB.Form frmMapa 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8610
   ClientLeft      =   7440
   ClientTop       =   3495
   ClientWidth     =   8610
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "FrizQuadrata BT"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa de Abraxas, apretá Arriba y Abajo para cambiar entre mapa global / laberintos. Apretá cualquier otra tecla para salir."
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   8040
      Width           =   8175
   End
   Begin VB.Image imgMap 
      Appearance      =   0  'Flat
      Height          =   7920
      Left            =   0
      Top             =   0
      Width           =   8595
   End
   Begin VB.Image imgMapDungeon 
      Appearance      =   0  'Flat
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp 'Cambiamos el "nivel" del mapa, al estilo Zelda ;D
            ToggleImgMaps
        Case Else
            Unload Me
    End Select
    
End Sub

Private Sub ToggleImgMaps()
    imgMap.Visible = Not imgMap.Visible
    imgMapDungeon.Visible = Not imgMapDungeon.Visible
End Sub

'Load the images. Resizes the form, adjusts image's left and top and set lblTexto's Top and Left.
Private Sub Form_Load()

On Error GoTo Error
    Dim x As Long
    Dim y As Long
    Dim n As Long

    x = Width / Screen.TwipsPerPixelX
    y = Height / Screen.TwipsPerPixelY

    'set the corner angle by changing the value of 'n'
    n = 25

    SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, x, y, n, n), True

    'Cargamos las imagenes de los mapas
    imgMap.Picture = LoadPicture(GrhPath & "mapa1.jpg")
    imgMapDungeon.Picture = LoadPicture(GrhPath & "mapa2.jpg")
    
    
    'Ajustamos el tamaño del formulario a la imagen más grande
    If imgMap.Width > imgMapDungeon.Width Then
        Width = imgMap.Width
    Else
        Width = imgMapDungeon.Width
    End If
    
    If imgMap.Height > imgMapDungeon.Height Then
        Height = imgMap.Height + lblTexto.Height
    Else
        Height = imgMapDungeon.Height + lblTexto.Height
    End If
    
    'Movemos ambas imágenes al centro del formulario
    imgMap.Left = Width * 0.5 - imgMap.Width * 0.5
    imgMap.Top = (Height - lblTexto.Height) * 0.5 - imgMap.Height * 0.5
    
    imgMapDungeon.Left = Width * 0.5 - imgMapDungeon.Width * 0.5
    imgMapDungeon.Top = (Height - lblTexto.Height) * 0.5 - imgMapDungeon.Height * 0.5
    
    lblTexto.Top = Height - lblTexto.Height
    lblTexto.Left = Width * 0.5 - lblTexto.Width * 0.5
    
    imgMapDungeon.Visible = False
    
    Call Make_Transparent_Form(hWnd, 200)
    Exit Sub
Error:
    MsgBox Err.Description, vbInformation, "Error: " & Err.Number
    Unload Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Unload Me
    End If
End Sub

Private Sub imgMap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    Unload Me
End If
End Sub

Private Sub imgMapDungeon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    Unload Me
End If
End Sub

Private Sub lblTexto_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    Unload Me
End If
End Sub
