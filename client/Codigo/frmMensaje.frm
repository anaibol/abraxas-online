VERSION 5.00
Begin VB.Form frmMensaje 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "FrizQuadrata BT"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   720
      Tag             =   "1"
      Top             =   2685
      Width           =   2655
   End
   Begin VB.Label msg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    SetFocus
End Sub

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Picture = LoadPicture(GrhPath & "VentanaMsj.jpg")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarMsj.jpg", _
                                    GrhPath & "BotónCerrarRolloverMsj.jpg", _
                                    GrhPath & "BotónCerrarClickMsj.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub msg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub
