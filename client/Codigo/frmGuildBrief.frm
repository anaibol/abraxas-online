VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7455
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
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "&H8000000A&"
   Begin VB.TextBox Desc 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2910
      Width           =   6930
   End
   Begin VB.Label antifaccion 
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
      Left            =   2280
      TabIndex        =   10
      Top             =   2505
      Width           =   2415
   End
   Begin VB.Label Aliados 
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
      Left            =   1800
      TabIndex        =   9
      Top             =   2175
      Width           =   1575
   End
   Begin VB.Label Enemigos 
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
      Left            =   1920
      TabIndex        =   8
      Top             =   1845
      Width           =   2175
   End
   Begin VB.Label lblAlineacion 
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
      Left            =   5280
      TabIndex        =   7
      Top             =   1860
      Width           =   1815
   End
   Begin VB.Label eleccion 
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
      Left            =   5280
      TabIndex        =   6
      Top             =   1155
      Width           =   1815
   End
   Begin VB.Label Miembros 
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
      Left            =   5160
      TabIndex        =   5
      Top             =   1500
      Width           =   1935
   End
   Begin VB.Label lider 
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
      Left            =   1080
      TabIndex        =   4
      Top             =   1140
      Width           =   3135
   End
   Begin VB.Label Creacion 
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
      Left            =   5760
      TabIndex        =   3
      Top             =   780
      Width           =   1455
   End
   Begin VB.Label fundador 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   810
      Width           =   2775
   End
   Begin VB.Label nombre 
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
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   120
      Tag             =   "1"
      Top             =   3990
      Width           =   1455
   End
   Begin VB.Image imgOfrecerPaz 
      Height          =   375
      Left            =   1680
      Tag             =   "1"
      Top             =   3990
      Width           =   1335
   End
   Begin VB.Image imgOfrecerAlianza 
      Height          =   375
      Left            =   3120
      Tag             =   "1"
      Top             =   3990
      Width           =   1335
   End
   Begin VB.Image imgDeclararGuerra 
      Height          =   375
      Left            =   4560
      Tag             =   "1"
      Top             =   3990
      Width           =   1335
   End
   Begin VB.Image imgSolicitarIngreso 
      Height          =   375
      Left            =   6000
      Tag             =   "1"
      Top             =   3990
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonGuerra As clsGraphicalButton
Private cBotonAlianza As clsGraphicalButton
Private cBotonPaz As clsGraphicalButton
Private cBotonSolicitarIngreso As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public EsLeader As Boolean

Private Sub Desc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    Call LoadButtons
    
End Sub

Private Sub LoadButtons()

    Set cBotonGuerra = New clsGraphicalButton
    Set cBotonAlianza = New clsGraphicalButton
    Set cBotonPaz = New clsGraphicalButton
    Set cBotonSolicitarIngreso = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonGuerra.Initialize(imgDeclararGuerra, GrhPath & "BotónDeclararGuerra.jpg", _
                                    GrhPath & "BotónDeclararGuerraRollover.jpg", _
                                    GrhPath & "BotónDeclararGuerraClick.jpg", Me)

    Call cBotonAlianza.Initialize(imgOfrecerAlianza, GrhPath & "BotónOfrecerAlianza.jpg", _
                                    GrhPath & "BotónOfrecerAlianzaRollover.jpg", _
                                    GrhPath & "BotónOfrecerAlianzaClick.jpg", Me)

    Call cBotonPaz.Initialize(imgOfrecerPaz, GrhPath & "BotónOfrecerPaz.jpg", _
                                    GrhPath & "BotónOfrecerPazRollover.jpg", _
                                    GrhPath & "BotónOfrecerPazClick.jpg", Me)

    Call cBotonSolicitarIngreso.Initialize(imgSolicitarIngreso, GrhPath & "BotónSolicitarIngreso.jpg", _
                                    GrhPath & "BotónSolicitarIngresoRollover.jpg", _
                                    GrhPath & "BotónSolicitarIngresoClick.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarDetallesGuilda.jpg", _
                                    GrhPath & "BotónCerrarRolloverDetallesGuilda.jpg", _
                                    GrhPath & "BotónCerrarClickDetallesGuilda.jpg", Me)


    If Not EsLeader Then
        imgDeclararGuerra.Visible = False
        imgOfrecerAlianza.Visible = False
        imgOfrecerPaz.Visible = False
    End If

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

Private Sub imgDeclararGuerra_Click()
    Call WriteGuildDeclareWar(Nombre.Caption)
    Unload Me
End Sub

Private Sub imgOfrecerAlianza_Click()
    frmCommet.Nombre = Nombre.Caption
    frmCommet.T = TIPO.ALIANZA
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub imgOfrecerPaz_Click()
    frmCommet.Nombre = Nombre.Caption
    frmCommet.T = TIPO.PAZ
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub imgSolicitarIngreso_Click()
    Call frmGuildSol.RecieveSolicitud(Nombre.Caption)
    Call frmGuildSol.Show(vbModal, frmGuildBrief)
End Sub

