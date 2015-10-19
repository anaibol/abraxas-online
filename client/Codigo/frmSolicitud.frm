VERSION 5.00
Begin VB.Form frmGuildSol 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4680
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
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1035
      Left            =   300
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Image imgEnviar 
      Height          =   285
      Left            =   1560
      Tag             =   "1"
      Top             =   2760
      Width           =   1185
   End
   Begin VB.Image imgCerrar 
      Height          =   285
      Left            =   240
      Tag             =   "1"
      Top             =   2760
      Width           =   1305
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Private cBotonEnviar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Dim CName As String

Public Sub RecieveSolicitud(ByVal GuildName As String)

    CName = GuildName

End Sub

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Picture = LoadPicture(GrhPath & "SolicitudIngresoGuilda.jpg")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonEnviar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarIngreso.jpg", _
                                    GrhPath & "BotónCerrarRolloverIngreso.jpg", _
                                    GrhPath & "BotónCerrarClickIngreso.jpg", Me)

    Call cBotonEnviar.Initialize(imgEnviar, GrhPath & "BotónEnviarIngreso.jpg", _
                                    GrhPath & "BotónEnviarRolloverIngreso.jpg", _
                                    GrhPath & "BotónEnviarClickIngreso.jpg", Me)
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

Private Sub imgEnviar_Click()
    Call WriteGuildRequestMembership(CName, Replace(Replace(Text1.Text, ",", ";"), vbCrLf, "º"))

    Unload Me
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub
