VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Frag Shooter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   960
      Left            =   2295
      TabIndex        =   17
      Top             =   3990
      Width           =   4380
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3825
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "5"
         Top             =   570
         Width           =   450
      End
      Begin VB.CheckBox chkFragShooter 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Al morir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1125
      End
      Begin VB.CheckBox chkFragShooter 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Al matar personajes mayores a nivel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3825
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Noticias del clan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   975
      Left            =   2295
      TabIndex        =   14
      Top             =   2925
      Width           =   4380
      Begin VB.OptionButton optMostrarNoticias 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Mostrar noticias al conectarse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   3240
      End
      Begin VB.OptionButton optNoMostrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "No mostrarlas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdChangePassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Cambiar Contraseña"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   13
      Top             =   900
      UseMaskColor    =   -1  'True
      Width           =   1900
   End
   Begin VB.CommandButton cmdCustomKeys 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Configurar Teclas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   1980
      UseMaskColor    =   -1  'True
      Width           =   1900
   End
   Begin VB.CommandButton customMsgCmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Mensajes Personalizados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   3060
      UseMaskColor    =   -1  'True
      Width           =   1900
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Audio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1020
      Left            =   2295
      TabIndex        =   7
      Top             =   690
      Width           =   4380
      Begin VB.CheckBox ChckSndFX 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Efectos de sonido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1875
      End
      Begin MSComctlLib.Slider SlMsc 
         Height          =   255
         Left            =   2200
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.CheckBox ChckMsc 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Música"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.Slider SlSndFx 
         Height          =   255
         Left            =   2200
         TabIndex        =   10
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   10
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Diálogos de clan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   975
      Left            =   2295
      TabIndex        =   2
      Top             =   1815
      Width           =   4380
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1515
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "5"
         Top             =   315
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "En pantalla,"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optConsola 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "En consola"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1995
         TabIndex        =   6
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Guardar y Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H00000000&
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4140
      UseMaskColor    =   -1  'True
      Width           =   1900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   600
      Left            =   1920
      TabIndex        =   1
      Top             =   135
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loading As Boolean

Private Sub chckMsc_Click()
    Exit Sub
    
    If ChckMsc.Value = vbChecked Then
        MusicActivated = True
        SlMsc.Enabled = True
    Else
        MusicActivated = False
        SlMsc.Enabled = False
        Call Audio.MusicMP3Stop
        Call Audio.mMusic_StopMid
    End If
    
    If Not loading Then
        Call Audio.mSound_PlayWav(SND_CLICK)
    End If
End Sub

Private Sub chckSndFx_Click()
    If ChckSndFX.Value = vbChecked Then
        SoundEffectsActivated = True
        SlSndFx.Enabled = True
    Else
        SoundEffectsActivated = False
        SlSndFx.Enabled = False
    End If
    
    If Not loading Then
        Call Audio.mSound_PlayWav(SND_CLICK)
    End If
End Sub

Private Sub cmdCustomKeys_Click()
    If Not loading Then
        Call Audio.mSound_PlayWav(SND_CLICK)
    End If
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub cmdChangePassword_Click()
    Call frmChangePassword.Show(vbModal, Me)
End Sub

Private Sub cmdViewMap_Click()
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub customMsgCmd_Click()
    Call frmMessageTxt.Show(vbModeless, Me)
End Sub

Private Sub Form_Load()
    
    Dim X As Long
    Dim Y As Long
    Dim n As Long

    X = Width / Screen.TwipsPerPixelX
    Y = Height / Screen.TwipsPerPixelY

    'set the corner angle by changing the value of 'n'
    n = 25

    SetWindowRgn hwnd, CreateRoundRectRgn(0, 0, X, Y, n, n), True

    loading = True      'Prevent sounds when setting check's values
    
    If MusicActivated Then
        ChckMsc.Value = vbChecked
        SlMsc.Enabled = True
        'SlMsc.value = Music_Volume
    Else
        ChckMsc.Value = vbUnchecked
    End If
    
    If SoundEffectsActivated Then
        ChckSndFX.Value = vbChecked
        SlSndFx.Enabled = True
        'SlSndFx.value = SoundVolume
    Else
        ChckSndFX.Value = vbUnchecked
        SlSndFx.Enabled = False
    End If
    
    'If ClientSetup.GuildNews Then
    'optMostrarNoticias.value = True
    'optNoMostrar.value = False
    'Else
        optMostrarNoticias.Value = False
        optNoMostrar.Value = True
    'End If
    
    Call Make_Transparent_Form(hwnd, 200)
    
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call Auto_Drag(hwnd)
Else
    Unload Me
End If
End Sub

Private Sub optMostrarNoticias_Click()
    'ClientSetup.GuildNews = True
End Sub

Private Sub optNoMostrar_Click()
    'ClientSetup.GuildNews = False
End Sub

'PRIVATE SUB Slider1_Change(Index As Integer)
'Select Case Index
'Case 0
'Audio.MusicVolume = Slider1(0).value
'Case 1
'Audio.SoundVolume = Slider1(1).value
'End Select
'END SUB

'PRIVATE SUB Slider1_Scroll(Index As Integer)
'Select Case Index
'Case 0
'Audio.MusicVolume = Slider1(0).value
'Case 1
'Audio.SoundVolume = Slider1(1).value
'End Select
'END SUB
