VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   0  'None
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6390
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
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPeticiones 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3210
      Width           =   5730
   End
   Begin VB.TextBox txtMiembro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4695
      Width           =   5730
   End
   Begin VB.Label Muertes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Image imgAceptar 
      Height          =   510
      Left            =   5160
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgRechazar 
      Height          =   510
      Left            =   3840
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgPeticion 
      Height          =   510
      Left            =   2640
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgEchar 
      Height          =   510
      Left            =   1440
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgCerrar 
      Height          =   510
      Left            =   120
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   700
      Width           =   1440
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1750
      Width           =   1185
   End
   Begin VB.Label Clase 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1225
      Width           =   1575
   End
   Begin VB.Label Raza 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label Genero 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label Banco 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2250
      Width           =   1425
   End
   Begin VB.Label guildactual 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   2265
   End
   Begin VB.Label Matados 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4905
      TabIndex        =   2
      Top             =   1500
      Width           =   1185
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Private cBotonPeticion As clsGraphicalButton
Private cBotonRechazar As clsGraphicalButton
Private cBotonEchar As clsGraphicalButton
Private cBotonAceptar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public Enum CharInfoFrmType
    frmMembers
    frmMembershipRequests
End Enum

Public frmType As CharInfoFrmType

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Picture = LoadPicture(GrhPath & "VentanaInfoPj.jpg")
    
    Call LoadButtons
    
End Sub

Private Sub LoadButtons()

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonPeticion = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    Set cBotonEchar = New clsGraphicalButton
    Set cBotonAceptar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarInfoChar.jpg", _
                                    GrhPath & "BotónCerrarRolloverInfoChar.jpg", _
                                    GrhPath & "BotónCerrarClickInfoChar.jpg", Me)

    Call cBotonPeticion.Initialize(imgPeticion, GrhPath & "BotónPetición.jpg", _
                                    GrhPath & "BotónPeticiónRollover.jpg", _
                                    GrhPath & "BotónPeticiónClick.jpg", Me)

    Call cBotonRechazar.Initialize(imgRechazar, GrhPath & "BotónRechazar.jpg", _
                                    GrhPath & "BotónRechazarRollover.jpg", _
                                    GrhPath & "BotónRechazarClick.jpg", Me)

    Call cBotonEchar.Initialize(imgEchar, GrhPath & "BotónEchar.jpg", _
                                    GrhPath & "BotónEcharRollover.jpg", _
                                    GrhPath & "BotónEcharClick.jpg", Me)
                                    
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotónAceptarInfoChar.jpg", _
                                    GrhPath & "BotónAceptarRolloverInfoChar.jpg", _
                                    GrhPath & "BotónAceptarClickInfoChar.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgAceptar_Click()
    Call WriteGuildAcceptNewMember(Nombre)
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub imgEchar_Click()
    Call WriteGuildKickMember(Nombre)
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub imgPeticion_Click()
    Call WriteGuildRequestJoinerInfo(Nombre)
End Sub

Private Sub imgRechazar_Click()
    frmCommet.T = rechazóPJ
    frmCommet.Nombre = Nombre.Caption
    frmCommet.Show vbModeless, frmCharInfo
End Sub

Private Sub txtMiembro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub
