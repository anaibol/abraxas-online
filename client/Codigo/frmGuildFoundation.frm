VERSION 5.00
Begin VB.Form frmGuildFoundation 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtGuildName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   3345
   End
   Begin VB.Image imgSiguiente 
      Height          =   375
      Left            =   2400
      Tag             =   "1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Image imgCancelar 
      Height          =   375
      Left            =   360
      Tag             =   "1"
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonSiguiente As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    SetFocus
End Sub

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Picture = LoadPicture(GrhPath & "VentanaNombreGuilda.jpg")
        
    Call LoadButtons
    
    If Len(txtGuildName.Text) <= 30 Then
        If Not AsciiValidos(txtGuildName) Then
            MsgBox "Nombre invalido."
            Exit Sub
        End If
    Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub
    End If

End Sub

Private Sub LoadButtons()

    Set cBotonSiguiente = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonSiguiente.Initialize(imgSiguiente, GrhPath & "BotónSiguienteNombreGuilda.jpg", _
                                    GrhPath & "BotónSiguienteRolloverNombreGuilda.jpg", _
                                    GrhPath & "BotónSiguienteClickNombreGuilda.jpg", Me)

    Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "BotónCancelarNombreGuilda.jpg", _
                                    GrhPath & "BotónCancelarRolloverNombreGuilda.jpg", _
                                    GrhPath & "BotónCancelarClickNombreGuilda.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub imgSiguiente_Click()
    GuildName = txtGuildName.Text
    Unload Me
    frmGuildDetails.Show , frmMain
End Sub

Private Sub txtGuildName_Change()
    txtGuildName.Text = StrConv(txtGuildName.Text, vbProperCase)
End Sub
