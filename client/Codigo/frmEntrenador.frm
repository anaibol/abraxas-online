VERSION 5.00
Begin VB.Form frmEntrenador 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   930
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "¿Con qué criatura querés luchar?"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   2865
   End
End
Attribute VB_Name = "frmEntrenador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lstCriaturasClick As Boolean

Private Sub Form_Load()

    Dim x As Long
    Dim y As Long
    Dim n As Long
    
    x = Width / Screen.TwipsPerPixelX
    y = Height / Screen.TwipsPerPixelY
    
    'set the corner angle by changing the value of 'n'
    n = 25
    
    SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, x, y, n, n), True
    
    Call Make_Transparent_Form(hWnd, 200)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call Auto_Drag(hWnd)
    Else
        Unload Me
    End If
End Sub

Private Sub lstCriaturas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lstCriaturasClick = True
End Sub

Private Sub lstCriaturas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lstCriaturasClick Then
        Exit Sub
    End If
    
    If Button = vbRightButton Then
        Unload Me
        lstCriaturasClick = False
    Else
        Call WriteTrain(lstCriaturas.ListIndex + 1)
        lstCriaturasClick = False
        Unload Me
    End If
End Sub
