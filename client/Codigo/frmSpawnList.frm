VERSION 5.00
Begin VB.Form frmSpawnList 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   2685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "FrizQuadrata BT"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   2730
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   2505
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NPC's"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   765
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call Auto_Drag(hWnd)
    Else
        Unload Me
    End If
End Sub

Private Sub lstCriaturas_DblClick()
    Call WriteSpawnCreature(lstCriaturas.ListIndex + 1)
End Sub

Private Sub lstCriaturas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    ElseIf KeyAscii = vbKeyReturn And lstCriaturas.ListIndex + 1 > 0 Then
        Call WriteSpawnCreature(lstCriaturas.ListIndex + 1)
        Unload Me
    End If
End Sub

Private Sub lstCriaturas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Unload Me
    End If
End Sub
