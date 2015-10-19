VERSION 5.00
Begin VB.Form frmGuildDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6840
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
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1500
      Left            =   405
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   6015
   End
   Begin VB.Image imgConfirmar 
      Height          =   360
      Left            =   4920
      Tag             =   "1"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   360
      Tag             =   "1"
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonConfirmar As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Const MAX_DESC_LENGTH As Integer = 520

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonConfirmar = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    'Call cBotonConfirmar.Initialize(imgConfirmar, GrhPath & "BotónConfirmarDesc.jpg", _
                                    GrhPath & "BotónConfirmarRolloverDesc.jpg", _
                                    GrhPath & "BotónConfirmarClickDesc.jpg", Me)

    'Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotónSalirDesc.jpg", _
                                    GrhPath & "BotónSalirRolloverDesc.jpg", _
                                    GrhPath & "BotónSalirClickDesc.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgConfirmar_Click()
    Dim Desc As String
    
    Desc = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)

    If CreandoGuilda Then
        Call WriteCreateNewGuild(Desc, GuildName)
    Else
        Call WriteGuildDescUpdate(Desc)
    End If

    CreandoGuilda = False
    Unload Me
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub txtDesc_Change()
    If Len(txtDesc.Text) > MAX_DESC_LENGTH Then
        txtDesc.Text = Left$(txtDesc.Text, MAX_DESC_LENGTH)
    End If
End Sub
