VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInvOroProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3450
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   930
      Width           =   960
   End
   Begin VB.TextBox txtAgregar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   270
      Left            =   4500
      TabIndex        =   6
      Top             =   2295
      Width           =   1035
   End
   Begin VB.PictureBox picInvOroOfertaOtro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   5040
      Width           =   960
   End
   Begin VB.PictureBox picInvOfertaOtro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   6975
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   5040
      Width           =   2400
   End
   Begin VB.PictureBox picInvOfertaProp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   6960
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   3
      Top             =   930
      Width           =   2400
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   495
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   7965
      Width           =   6060
   End
   Begin VB.PictureBox picInvComercio 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   630
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   945
      Width           =   2400
   End
   Begin VB.PictureBox picInvOroOfertaProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   930
      Width           =   960
   End
   Begin RichTextLib.RichTextBox CommerceConsole 
      Height          =   1620
      Left            =   495
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   6030
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   2858
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmComerciarUsu.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgCancelar 
      Height          =   360
      Left            =   480
      Tag             =   "1"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image imgRechazar 
      Height          =   360
      Left            =   8220
      Tag             =   "2"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image imgConfirmar 
      Height          =   360
      Left            =   7440
      Tag             =   "2"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image imgAceptar 
      Height          =   360
      Left            =   6750
      Tag             =   "2"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image imgAgregar 
      Height          =   255
      Left            =   4920
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgQuitar 
      Height          =   255
      Left            =   4920
      Top             =   2760
      Width           =   255
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
