VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Abraxas"
   ClientHeight    =   9000
   ClientLeft      =   -1170
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   8520
      Top             =   1200
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   7000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   6240
      Left            =   120
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   42
      Top             =   2280
      Width           =   8160
   End
   Begin VB.PictureBox picShip 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10560
      Picture         =   "frmMain.frx":4282
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   41
      Top             =   6120
      Width           =   480
   End
   Begin VB.PictureBox picBeltEqp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11520
      Picture         =   "frmMain.frx":71EA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   37
      Top             =   6120
      Width           =   480
   End
   Begin VB.PictureBox picRingEqp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10920
      Picture         =   "frmMain.frx":A152
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   36
      Top             =   6120
      Width           =   480
   End
   Begin VB.PictureBox PicBelt 
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
      Height          =   480
      Left            =   9300
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":D0BA
      ScaleHeight     =   30.567
      ScaleMode       =   0  'User
      ScaleWidth      =   120.47
      TabIndex        =   35
      Tag             =   "3"
      Top             =   2850
      Width           =   1920
   End
   Begin MSComctlLib.ImageList MouseImage 
      Left            =   11430
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.PictureBox picCompasScrollUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   8100
      Picture         =   "frmMain.frx":1FCFE
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   34
      Top             =   105
      Width           =   210
   End
   Begin VB.PictureBox picCompasScrollDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   8100
      Picture         =   "frmMain.frx":22993
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   33
      Top             =   1590
      Width           =   210
   End
   Begin VB.PictureBox picConsoleScrollDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6150
      Picture         =   "frmMain.frx":25620
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   31
      Top             =   1590
      Width           =   210
   End
   Begin VB.PictureBox picConsoleScrollUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6150
      Picture         =   "frmMain.frx":282AD
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   30
      Top             =   105
      Width           =   210
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9000
      Top             =   0
   End
   Begin RichTextLib.RichTextBox SendTxt 
      Height          =   330
      Left            =   120
      TabIndex        =   28
      Top             =   1920
      Visible         =   0   'False
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   582
      _Version        =   393217
      BackColor       =   128
      BorderStyle     =   0
      MultiLine       =   0   'False
      DisableNoScroll =   -1  'True
      MaxLength       =   100
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":2AF42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picLeftHandEqp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9600
      Picture         =   "frmMain.frx":2AFB9
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   6045
      Width           =   480
   End
   Begin VB.PictureBox picHeadEqp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8640
      Picture         =   "frmMain.frx":2DFC0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   6045
      Width           =   480
   End
   Begin VB.PictureBox picRightHandEqp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10080
      Picture         =   "frmMain.frx":30EC2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   6045
      Width           =   480
   End
   Begin VB.PictureBox picBodyEqp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9120
      Picture         =   "frmMain.frx":33E2A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   6120
      Width           =   480
   End
   Begin VB.Timer Macro 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   8520
      Top             =   840
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   8520
      Top             =   480
   End
   Begin VB.Timer MacroTrabajo 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8520
      Top             =   45
   End
   Begin VB.PictureBox Minimap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   3  'Vertical Line
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1800
      Left            =   8520
      ScaleHeight     =   78.857
      ScaleMode       =   0  'User
      ScaleWidth      =   120.96
      TabIndex        =   0
      Top             =   240
      Width           =   1800
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H0080FFFF&
         Height          =   225
         Left            =   744
         Top             =   765
         Width           =   270
      End
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   120
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   128
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      DisableNoScroll =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":36D47
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox CompaRecTxt 
      Height          =   1740
      Left            =   6450
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   120
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   3069
      _Version        =   393217
      BackColor       =   128
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      DisableNoScroll =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":36DCB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicCompaInv 
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
      Height          =   1920
      Left            =   6390
      ScaleHeight     =   122.269
      ScaleMode       =   0  'User
      ScaleWidth      =   120.47
      TabIndex        =   27
      Tag             =   "5"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox PicInv 
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
      Height          =   1920
      Left            =   9060
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":36E4F
      ScaleHeight     =   122.269
      ScaleMode       =   0  'User
      ScaleWidth      =   150.588
      TabIndex        =   5
      Tag             =   "1"
      Top             =   3480
      Width           =   2400
   End
   Begin VB.ListBox lstSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2295
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":49A93
      Left            =   9240
      List            =   "frmMain.frx":49A9A
      TabIndex        =   29
      Top             =   3000
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.PictureBox PicSpellInv 
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
      Height          =   1920
      Left            =   9060
      Picture         =   "frmMain.frx":49AA9
      ScaleHeight     =   122.269
      ScaleMode       =   0  'User
      ScaleWidth      =   150.588
      TabIndex        =   7
      Tag             =   "2"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Image imgInv 
      Appearance      =   0  'Flat
      Height          =   2970
      Left            =   8760
      Top             =   2760
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label lblShip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   40
      Top             =   7920
      Width           =   675
   End
   Begin VB.Label lblRingEqp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   39
      Top             =   7560
      Width           =   675
   End
   Begin VB.Label lblBeltEqp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   38
      Top             =   7320
      Width           =   675
   End
   Begin VB.Image ImgSend 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7980
      MouseIcon       =   "frmMain.frx":5C6ED
      MousePointer    =   99  'Custom
      Top             =   1875
      Width           =   345
   End
   Begin VB.Label lblHechizos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   10320
      MouseIcon       =   "frmMain.frx":5C83F
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label lblInventario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   9000
      MouseIcon       =   "frmMain.frx":5C991
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Image ImgMascotas 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   10425
      MouseIcon       =   "frmMain.frx":5CAE3
      MousePointer    =   99  'Custom
      Top             =   8460
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image ImgEstadisticas 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   9975
      MouseIcon       =   "frmMain.frx":5CC35
      MousePointer    =   99  'Custom
      Top             =   8445
      Width           =   480
   End
   Begin VB.Shape SendTxtBorder 
      BorderColor     =   &H0080FFFF&
      Height          =   285
      Left            =   120
      Top             =   1890
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Label COMIDALbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "68%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   180
      Left            =   9600
      TabIndex        =   23
      ToolTipText     =   "Hambre"
      Top             =   7965
      Width           =   465
   End
   Begin VB.Label AGUALbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "68%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   180
      Left            =   9150
      TabIndex        =   22
      ToolTipText     =   "Sed"
      Top             =   7965
      Width           =   465
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   21
      Top             =   8640
      Width           =   8145
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   9960
      TabIndex        =   20
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label GldLbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   210
      Left            =   10950
      TabIndex        =   18
      Top             =   7605
      Width           =   105
   End
   Begin VB.Label lblDext 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   210
      Left            =   10950
      TabIndex        =   15
      ToolTipText     =   "Agilidad"
      Top             =   7245
      Width           =   330
   End
   Begin VB.Label lblStrg 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   10950
      TabIndex        =   14
      ToolTipText     =   "Fuerza"
      Top             =   6885
      Width           =   330
   End
   Begin VB.Label LblPoblacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   10950
      TabIndex        =   13
      ToolTipText     =   "Población actual de Abraxas. Hacé click para ver los nombres de los pobladores."
      Top             =   7965
      Width           =   465
   End
   Begin VB.Image imgDropGold 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   10530
      Top             =   7575
      Width           =   360
   End
   Begin VB.Label lblHeadEqp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   11
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label lblLeftHandEqp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   6600
      Width           =   675
   End
   Begin VB.Label lblBodyEqp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      Top             =   6360
      Width           =   675
   End
   Begin VB.Label lblRightHandEqp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   6600
      Width           =   675
   End
   Begin VB.Image picListaCompa 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8520
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image picInvCompa 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8535
      Top             =   1620
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000040&
      Height          =   6240
      Left            =   123
      Top             =   2318
      Width           =   8190
   End
   Begin VB.Label HPLbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0 / 100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   180
      Left            =   9000
      TabIndex        =   24
      Top             =   6885
      Width           =   1320
   End
   Begin VB.Label MANLbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0 / 100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   180
      Left            =   9000
      TabIndex        =   25
      Top             =   7245
      Width           =   1320
   End
   Begin VB.Image ImgMana 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   9000
      Picture         =   "frmMain.frx":5CD87
      Top             =   7290
      Width           =   1320
   End
   Begin VB.Image ImgHP 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   9000
      Picture         =   "frmMain.frx":5D60B
      Top             =   6930
      Width           =   1320
   End
   Begin VB.Image ImgGuildas 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   9525
      MouseIcon       =   "frmMain.frx":5DE8F
      MousePointer    =   99  'Custom
      Top             =   8445
      Width           =   435
   End
   Begin VB.Image ImgOpciones 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   9045
      MouseIcon       =   "frmMain.frx":5DFE1
      MousePointer    =   99  'Custom
      Top             =   8460
      Width           =   435
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   255
      Left            =   11400
      MousePointer    =   99  'Custom
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label STALbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0 / 100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   180
      Left            =   9000
      TabIndex        =   26
      Top             =   7605
      Width           =   1320
   End
   Begin VB.Image ImgSta 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   9000
      Picture         =   "frmMain.frx":5E133
      Top             =   7650
      Width           =   1320
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   11400
      TabIndex        =   19
      Top             =   2400
      Width           =   300
   End
   Begin VB.Label ExpLbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   180
      Left            =   8880
      TabIndex        =   12
      Top             =   2415
      Width           =   2535
   End
   Begin VB.Image ImgExp 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   8850
      Picture         =   "frmMain.frx":5E9B7
      Top             =   2460
      Width           =   2565
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Private clicX As Long
Private clicY As Long

Private ButtonClicked As Integer

Public IsPlaying As Byte

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim DoSReady As Boolean

Public Sub ActivarMacroHechizos()

    If UserMuerto Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
        Exit Sub
    End If
    
    If SpellSelSlot < 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Primero seleccioná un hechizo.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
        Exit Sub
    End If
    
    TrainingMacro.Interval = INT_MACRO_HECHIS
    TrainingMacro.Enabled = True
    UsingSkill = Magia
    MousePointer = 2
        
    Call ShowConsoleMsg("Macro de hechizos activado.", 0, 200, 200, False, False, False)
End Sub

Public Sub DesactivarMacroHechizos()
    TrainingMacro.Enabled = False
    Call ShowConsoleMsg("Macro de hechizos desactivado.", 0, 150, 150, False, False, False)
End Sub

Private Sub CompaRecTxt_GotFocus()
    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        PicSpellInv.SetFocus
    End If
End Sub

Private Sub CompaRecTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim txt As String
    Dim Fa As Boolean
    Dim So As Byte
    
    txt = RichWordOver(CompaRecTxt, X, Y)
    
    If LenB(txt) > 2 Then
        So = EsCompaniero(txt)
        
        If So > 0 Then
            If Compa(So).Online Then
                Fa = True
            End If
        End If
    End If
    
    If Fa Then
        CompaRecTxt.MousePointer = rtfArrowQuestion
        CompaRecTxt.ToolTipText = "Hablar con " & txt
    Else
        If CompaRecTxt.MousePointer <> 1 Then
            CompaRecTxt.MousePointer = 1
            CompaRecTxt.ToolTipText = vbNullString
        End If
    End If
    
End Sub

Private Sub CompaRecTxt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error Resume Next
 
    Dim CompaName As String
    
    CompaName = RichWordOver(CompaRecTxt, X, Y)

    If LenB(CompaName) < 3 Then
        Exit Sub
    End If
    
    If Button = vbRightButton Then
        If Charlist(UserCharIndex).Priv > 1 Then
            If Compa(EsCompaniero(CompaName)).Online Then
                Call WriteGoToChar(CompaName)
                Exit Sub
            End If
        End If
    End If
        
    Dim i As Integer
    Dim a As Boolean
            
    If SendTxt.Visible Then
                        
        If LenB(SendTxt.Text) < 1 Then
            
            For i = 1 To LastChar
                If Charlist(i).EsUser Then
                    If LenB(Charlist(i).Nombre) = LenB(CompaName) Then
                        If Charlist(i).Nombre = CompaName Then
                            SelectedCharIndex = i
                            a = True
                            Exit For
                        End If
                    End If
                End If
            Next i
            
            If Not a Then
                If Not Compa(EsCompaniero(CompaName)).Online Then
                    Exit Sub
                End If
            End If
            
            If InStr(CompaName, " ") > 0 Then
                CompaName = Replace(CompaName, " ", "+")
            End If
        
            SendTxt.Text = ":" & CompaName & " "
            SendTxt.SetFocus
            SendTxt.SelStart = Len(SendTxt)
                            
        Else
        
            If InStr(CompaName, " ") > 0 Then
                CompaName = Replace(CompaName, " ", "+")
                a = True
            End If
            
            If Left$(SendTxt.Text, Len(CompaName) + 2) = ":" & CompaName & " " Then
                SendTxt.Visible = False
                SendTxt.Text = vbNullString
                
                If SelectedCharIndex > 0 Then
                    If Not Charlist(SelectedCharIndex).EsUser Then
                        SelectedCharIndex = 0
                    End If
                End If
            Else

                If a Then
                    CompaName = Replace(CompaName, "+", " ")
                End If
                
                a = False
                
                For i = 1 To LastChar
                    If Not Charlist(i).EsUser Then
                        If LenB(Charlist(i).Nombre) = LenB(CompaName) Then
                            If Charlist(i).Nombre = CompaName Then
                                SelectedCharIndex = i
                                a = True
                                Exit For
                            End If
                        End If
                    End If
                Next i
                
                If Not a Then
                    If Not Compa(EsCompaniero(CompaName)).Online Then
                        Exit Sub
                    End If
                End If
            
                SendTxt.Text = ":" & CompaName & " " & SendTxt.Text
                SendTxt.SetFocus
                SendTxt.SelStart = Len(SendTxt)
            End If
        
        End If
        
    Else
    
        For i = 1 To LastChar
            If Charlist(i).EsUser Then
                If Charlist(i).Nombre = CompaName Then
                    If LenB(Charlist(i).Nombre) = LenB(CompaName) Then
                        SelectedCharIndex = i
                        a = True
                        Exit For
                    End If
                End If
            End If
        Next i
        
        If Not a Then
            If Not Compa(EsCompaniero(CompaName)).Online Then
                Exit Sub
            End If
        End If
    
        If InStr(CompaName, " ") > 0 Then
            CompaName = Replace(CompaName, " ", "+")
        End If
            
        SendTxt.Text = ":" & CompaName & " "
        SendTxt.Visible = True
        SendTxt.SetFocus
        SendTxt.SelStart = Len(SendTxt)
        
    End If

End Sub

Private Sub ExpLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ExpLbl.Caption = UserExp & " / " & UserPasarNivel
End Sub

Private Sub Form_Activate()
    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next
    
    Select Case KeyCode
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            
            If Not frmComerciar.Visible And Not frmComerciarUsu.Visible And _
            Not frmBanco.Visible And Not frmSkills.Visible And _
            Not frmMSG.Visible And Not frmEntrenador.Visible And _
            Not frmEstadisticas.Visible And Not frmCantidad.Visible And Not frmCantidadGld.Visible Then
                
                If LenB(LastParsedString) > 3 Then
                    If Left$(LastParsedString, 1) = ":" Then
                        If LastParsedString <> ":P" And _
                            LastParsedString <> ":D" And _
                            LastParsedString <> ":)" And _
                            LastParsedString <> ":(" Then
        
                            SendTxt.Text = LastParsedString
                            SendTxt.SelStart = Len(SendTxt)
                        End If
                    End If
                End If
                
                SendTxt.Visible = True
                
                SendTxt.SetFocus
            End If
            
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
            KeyCode = 0
        
        Case vbKeyDelete, 110
            If Not SendTxt.Visible Then
                SendTxt.Text = "."
                SendTxt.Visible = True
                SendTxt.SetFocus
                SendTxt.SelStart = Len(SendTxt)
            End If

        Case vbKeyMultiply
            If Not SendTxt.Visible Then
                SendTxt.Text = "*"
                SendTxt.Visible = True
                SendTxt.SetFocus
                SendTxt.SelStart = Len(SendTxt)
            End If
            
        Case vbKeyDivide
            If Not SendTxt.Visible Then
                SendTxt.Text = "/"
                SendTxt.Visible = True
                SendTxt.SetFocus
                SendTxt.SelStart = Len(SendTxt)
            End If
            
        Case 186
            If Shift Then
                If Not SendTxt.Visible Then
                    SendTxt.Text = ":"
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                    SendTxt.SelStart = Len(SendTxt)
                End If
            End If
            
            'Else
            '    If Not SendTxt.Visible Then
            '        SendTxt.Text = "."
            '        SendTxt.Visible = True
            '        SendTxt.SetFocus
            '        SendTxt.SelStart = Len(SendTxt)
            '    End If
            'End If
            
        Case vbKeyEscape
            Call WriteQuit
            
        Case vbKeySpace
            If UserMuerto Then
                Call WriteHome
            End If
            
            'If Shift Then
            '    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
            
            'Else
            '    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
            '    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
            'End If
            
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call Form_KeyDown(KeyAscii, 0)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not SendTxt.Visible Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
            
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    
                    If Shift > 0 Then
                        If MusicActivated Then
                            Call Audio.mMusic_StopMid
                                            
                            Call Audio.MusicMP3Stop
                            
                            MusicActivated = False
                        Else
                            MusicActivated = True
                        End If
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                
                    If UserMuerto Then 'Muerto
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                        Exit Sub
                    End If
                    
                    If MapData(UserPos.X, UserPos.Y).Obj.Amount > 0 Then
                        If Not MapData(UserPos.X, UserPos.Y).Blocked Then
                            'If MapData(UserPos.X, UserPos.Y).TileExit.Map < 1 Then
                                If LenB(MapData(UserPos.X, UserPos.Y).Obj.Name) > 0 Then
                                    Call WritePickUp
                                End If
                            'End If
                        End If
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call Equipar
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    
                    If UserMuerto Then 'Muerto
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                        Exit Sub
                    End If
                    
                    If UserMinSTA < 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("No tenés suficiente energía.", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                        Exit Sub
                    End If

                    MousePointer = 2
                    UsingSkill = Domar

                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    
                    If UserMuerto Then 'Muerto
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                        Exit Sub
                    End If
                    
                    If UserMinSTA < 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("No tenés suficiente energía.", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                        Exit Sub
                    End If

                    MousePointer = 2
                    UsingSkill = Robar
        
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    
                    If UserMuerto Then 'Muerto
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                        Exit Sub
                    End If
                    
                    If Charlist(UserCharIndex).Invisible Then
                        Exit Sub
                    End If
                    
                    Call WriteWork(eSkill.Ocultarse)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    
                    If UserMuerto Then 'Muerto
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                        Exit Sub
                    End If
                    
                    If Not MainTimer.Check(TimersIndex.Drop) Then
                        Exit Sub
                    End If
                    
                    If (InvSelSlot > 0 And InvSelSlot <= MaxInvSlots) Then
                        If Inv(InvSelSlot).Amount = 1 Or Shift > 0 Then
                            Call WriteDrop(InvSelSlot, Inv(InvSelSlot).Amount)
                            Call Audio.mSound_PlayWav(SND_DROP)
                            
                        Else
                            If Inv(InvSelSlot).Amount > 1 Then
                                frmCantidad.Show , Me
                            End If
                        End If
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UseItemWithU
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        If Not UserMoving Then
                            Call WriteRequestPositionUpdate
                        End If
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    
                    Call WriteResuscitationToggle
            
                Case KeyCodeConstants.vbKeyI
                    If PicInv.Visible Then
                        Call Audio.mSound_PlayWav(SND_CLICK)

                        imgInv.Visible = False
                        PicSpellInv.Visible = False
                        
                        PicInv.Visible = True
                    End If
                
                Case KeyCodeConstants.vbKeyH
                    If Not imgInv.Visible Then
                        Call Audio.mSound_PlayWav(SND_CLICK)
                        
                        PicInv.Visible = False
                        imgInv.Visible = True
                        PicSpellInv.Visible = True
                    End If
            End Select
            
        Else
            'Select Case KeyCode
                'Custom messages!
            '    Case vbKey0 To vbKey9
            '        Dim CustomMessage As String
                    
            '        CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)
            '        If LenB(CustomMessage) > 0 Then
                        'No se pueden mandar mensajes personalizados de clan o a companiero!
            '            If UCase$(Left(CustomMessage, 5)) <> "/CMSG" And _
            '                Left(CustomMessage, 1) <> ":" Then
            '                Call ParseUserCommand(CustomMessage)
            '            End If
            '        End If
            'End Select
            
            Select Case KeyCode
                Case vbKey1
                    If Not MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Exit Sub
                    End If
                    
                    If Belt(1).ObjIndex < 1 Then
                        Exit Sub
                    End If
                        
                    Call WriteUseBeltItem(1)
                
                Case vbKey2
                    If Not MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Exit Sub
                    End If
                    
                    If Belt(2).ObjIndex < 1 Then
                        Exit Sub
                    End If
                        
                    Call WriteUseBeltItem(2)
                
                Case vbKey3
                    If Not MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Exit Sub
                    End If
                    
                    If Belt(3).ObjIndex < 1 Then
                        Exit Sub
                    End If
                        
                    Call WriteUseBeltItem(3)
                
                Case vbKey4
                    If Not MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Exit Sub
                    End If
                    
                    If Belt(4).ObjIndex < 1 Then
                        Exit Sub
                    End If
                        
                    Call WriteUseBeltItem(4)
                
            End Select
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)

        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture
        
        Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
            FPSFLAG = Not FPSFLAG
            
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, Me)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            'If Meditando Then
            '    Call DoMeditar
            'End If
            
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
            If TrainingMacro.Enabled Then
                DesactivarMacroHechizos
            Else
                ActivarMacroHechizos
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If MacroTrabajo.Enabled Then
                DesactivarMacroTrabajo
            Else
                ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
        
            If Shift > 0 Then
                Exit Sub
            End If
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then
                Exit Sub 'Check if arrows interval has finished.
            End If
            
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then
                    Exit Sub 'Corto intervalo Golpe-Hechizo
                End If
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or Descansando Or Meditando Then
                    Exit Sub
                End If
            End If
            
            If UserMuerto Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                End With
                Exit Sub
            End If
        
            If UserMinSTA < 10 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("No tenés suficiente energía.", .Red, .Green, .Blue, .Bold, .Italic)
                End With
                Exit Sub
            End If
        
            If TrainingMacro.Enabled Then
                DesactivarMacroHechizos
            End If
            
            If MacroTrabajo.Enabled Then
                DesactivarMacroTrabajo
            End If
        
            If LeftHandEqp.ObjIndex > 0 Then
                If LeftHandEqp.Proyectil Then
                    UsingSkill = Proyectiles
                    MousePointer = 2
                    Exit Sub
                End If
                
            ElseIf RightHandEqp.ObjIndex > 0 Then
            
                If RightHandEqp.ObjType = otInstrumento Then
                    Call WriteUseItem(InvSelSlot)
                    
                Else
                    Select Case LeftHandEqp.ObjIndex
                    
                        Case 389 'MARTILLO_HERRERO
                            UsingSkill = Herreria
                            MousePointer = 2
                                
                        Case 187 'PIQUETE_MINERO
                            UsingSkill = Minar
                            MousePointer = 2
                            
                        Case 138, 543 'CAÑA_PESCA O Red_PESCA
                            UsingSkill = Pescar
                            MousePointer = 2
                            
                        Case 127, 1005 'HACHA_LEÑADOR O HACHA_LEÑA_ELFICA
                            UsingSkill = Talar
                            MousePointer = 2
                            
                        Case 15, 198 'DAGA O SERRUCHO_CARPINTERO
                            Call WriteUseItem(InvSelSlot)
                    End Select
                End If
            End If
            
            Call WriteAttack
            
    End Select

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Comerciando Then
        Exit Sub
    End If

    If Cartel Then
        Cartel = False
    End If

    clicX = X
    clicY = Y
    
    If Not InGameArea() Then
        Exit Sub
    End If
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
            
    Select Case Button
    
        Case vbLeftButton
                    
            If SelectedCharIndex > 0 Then
                If Charlist(SelectedCharIndex).EsUser Or Charlist(SelectedCharIndex).MascoIndex > 0 Then
                    SelectedCharIndex = 0
                End If

                If SendTxt.Visible Then
                    If MapData(tX, tY).CharIndex > 0 Then
                        If Charlist(MapData(tX, tY).CharIndex).MascoIndex < 1 Then
                            If Left$(SendTxt.Text, 1) = ":" Then
                                SendTxt.Visible = False
                                SendTxt.Text = vbNullString
                            End If
                        End If
                    End If
                End If
            End If
            
            If UserMuerto Then
                Call WriteLeftClick(tX, tY)
            Else
                If UsingSkill = 0 Then
                    If tX = UserPos.X And tY = UserPos.Y Then
                        If MapData(tX, tY).Obj.Amount > 0 Then
                            If Not MapData(tX, tY).Blocked Then
                                'If MapData(tX, tY).TileExit.Map < 1 Then
                                    If LenB(MapData(tX, tY).Obj.Name) > 0 Then
                                        Call WritePickUp
                                    End If
                                'End If
                            End If
                        Else
                            Call WriteLeftClick(tX, tY)
                        End If
                                                
                    ElseIf MapData(tX, tY).CharIndex > 0 Then
                        
                        If Charlist(MapData(tX, tY).CharIndex).MascoIndex > 0 Then
                            If SelectedCharIndex <> MapData(tX, tY).CharIndex Then
                                Call WriteLeftClick(tX, tY)
                                SelectedCharIndex = MapData(tX, tY).CharIndex
                            End If
                        Else
                            Call WriteLeftClick(tX, tY)
                        End If
                        
                    ElseIf MapData(tX, tY).Obj.Amount > 0 Then
                        'If MapData(tX, tY).TileExit.Map < 1 Then
                            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                Call ShowConsoleMsg(MapData(tX, tY).Obj.Name, .Red, .Green, .Blue, .Bold, .Italic)
                            End With
                        'End If
                        
                        Call WriteLeftClick(tX, tY)
                        
                    Else
                        Call WriteLeftClick(tX, tY)
                    End If
                
                Else
                    If TrainingMacro.Enabled Then
                        DesactivarMacroHechizos
                    End If
                    
                    If MacroTrabajo.Enabled Then
                        DesactivarMacroTrabajo
                    End If
                            
                    Select Case UsingSkill
        
                        Case Proyectiles
                            If Not MainTimer.Check(TimersIndex.Arrows) Then
                                Exit Sub
                            End If
                            
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            
                        Case Robar, Domar, Pescar, Talar, Minar, Herreria
                            If Not MainTimer.Check(TimersIndex.Work) Then
                                Exit Sub
                            End If

                            Call WriteWorkLeftClick(tX, tY, UsingSkill)

                        Case Magia
                            If UsaMacro Then
                                CnTd = CnTd + 1
                                If CnTd = 3 Then
                                    Call WriteUseSpellMacro
                                    CnTd = 0
                                End If
                                UsaMacro = False
                            End If
                                
                            If MousePointer <> 2 Then
                                Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                            End If
                        
                            If Descansando Or Meditando Then
                                MousePointer = vbDefault
                                Exit Sub
                            End If
            
                            If Not MainTimer.Check(TimersIndex.Attack) Then 'Check if attack interval has finished.
                                If Not MainTimer.Check(TimersIndex.CastAttack) Or _
                                Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                    MousePointer = vbDefault
                                    Exit Sub
                                End If
                            End If
                            
                            Call WriteCastSpell(SpellSelSlot, tX, tY)
                    
                        Case Else 'Fundir minerales
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    End Select
        
                    UsingSkill = 0
                    MousePointer = vbDefault
                End If
            End If
            
        Case vbRightButton
            
            If Shift > 0 Then
                If Charlist(UserCharIndex).Priv > 1 Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
            
            If MapData(tX, tY).CharIndex > 0 Then
            
                Dim i As Integer
                Dim Name As String

                i = MapData(tX, tY).CharIndex
                
                If i = UserCharIndex Then
                    Exit Sub
                End If
                
                If Charlist(i).Lvl > 0 Then
                
                    If Charlist(i).MascoIndex > 0 Then
                        
                        If Charlist(i).Quieto Then
                            Call WritePetFollow(Charlist(i).MascoIndex)
                            Charlist(i).Quieto = False
                            
                            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                Call ShowConsoleMsg(Charlist(i).Nombre & " te seguirá.", .Red, .Green, .Blue, .Bold, .Italic)
                            End With
                        Else
                            Call WritePetStand(Charlist(i).MascoIndex)
                            Charlist(i).Quieto = True
                            
                            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                Call ShowConsoleMsg(Charlist(i).Nombre & " se quedará en su lugar.", .Red, .Green, .Blue, .Bold, .Italic)
                            End With
                        End If
                        
                        SelectedCharIndex = i
                    End If
                    
                ElseIf LenB(Charlist(i).Nombre) > 0 Then
                         'If i = SelectedCharIndex Then
                                                                          
                    Name = Charlist(i).Nombre
                    
                    If InStr(Name, " ") > 0 Then
                        Name = Replace(Name, " ", "+")
                    End If
                    
                    If SendTxt.Visible Then
                                        
                        If LenB(SendTxt.Text) < 1 Then
                            SelectedCharIndex = i
                            
                            SendTxt.Text = ":" & Name & " "
                            SendTxt.SetFocus
                            SendTxt.SelStart = Len(SendTxt)
    
                        ElseIf Left$(SendTxt.Text, Len(Name) + 2) <> ":" & Name & " " Then
                            SelectedCharIndex = i
                            SendTxt.Text = ":" & Name & " " & SendTxt.Text
                            SendTxt.SetFocus
                            SendTxt.SelStart = Len(SendTxt)
                        End If
                        
                    Else
                        SelectedCharIndex = i
                    
                        SendTxt.Text = ":" & Name & " "
                        SendTxt.Visible = True
                        SendTxt.SetFocus
                        SendTxt.SelStart = Len(SendTxt)
                    End If
                End If
            End If
        
        Case vbMiddleButton
            If Charlist(UserCharIndex).Priv > 1 Then
                Call WriteWarpChar("YO", UserMap, tX, tY)
            End If
            
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub ImgSend_Click()

    If SendTxt.Visible Then
        If LenB(Trim$(SendTxt.Text)) < 1 Then
            Call WriteDeleteChat
            Call Dialogos.RemoveDialog(UserCharIndex)
        Else
            Call ParseUserCommand(SendTxt.Text)
        End If
        
        SendTxt.Text = vbNullString
        SendTxt.Visible = False
        
        If SelectedCharIndex > 0 Then
            If Charlist(SelectedCharIndex).EsUser Then
                SelectedCharIndex = 0
            End If
        End If
        
        'RecTxt.SetFocus
    Else
        If Not frmComerciar.Visible And Not frmComerciarUsu.Visible And _
        Not frmBanco.Visible And Not frmSkills.Visible And _
        Not frmMSG.Visible And Not frmEntrenador.Visible And _
        Not frmEstadisticas.Visible And Not frmCantidad.Visible And Not frmCantidadGld.Visible Then
            
            If Left$(LastParsedString, 1) = ":" Then
                If LastParsedString <> ":P" And _
                    LastParsedString <> ":D" And _
                    LastParsedString <> ":)" And _
                    LastParsedString <> ":(" Then

                    SendTxt.Text = LastParsedString
                    SendTxt.SelStart = Len(SendTxt)
                End If
            End If
            
            SendTxt.Visible = True
            
            SendTxt.SetFocus
        End If
    End If
End Sub

Private Sub ImgEstadisticas_Click()
    LlegaronAtrib = False
    Call WriteRequestAttributes
    Call WriteRequestMiniStats
    Call FlushBuffer
    
    Do While Not LlegaronAtrib
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , Me
    LlegaronAtrib = False
End Sub

Private Sub ImgGuildas_Click()
    Call ShowConsoleMsg("Las guildas se encuentran deshabilitadas.")
    
    Exit Sub
    
    If frmGuildLeader.Visible Then
        Unload frmGuildLeader
    End If
    
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub ImgOpciones_Click()
    Call frmOpciones.Show(vbModeless, Me)
End Sub

Private Sub lblMapName_Click()
    Call frmMapa.Show
End Sub

Private Sub picBelt_DblClick()

    If ButtonClicked = vbRightButton Then
        Exit Sub
    End If
    
    If frmCarp.Visible Then
        Exit Sub
    End If
    
    If frmHerrero.Visible Then
        Exit Sub
    End If
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
        Exit Sub
    End If
    
    If Comerciando Then
        Exit Sub
    End If
    
    If BeltSelSlot < 1 Or BeltSelSlot > MaxBeltSlots Then
        Exit Sub
    End If
    
    If Belt(BeltSelSlot).ObjIndex < 1 Then
        Exit Sub
    End If
    
    If UserMuerto Then
        'If Inv(Slot).ObjType <> otBarco Then
        '    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        '        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        '    End With
            Exit Sub
        'End If
    End If
    
    If MacroTrabajo.Enabled Then
        DesactivarMacroTrabajo
    End If
    
    If TrainingMacro.Enabled Then
        DesactivarMacroHechizos
    End If
    
    Call WriteUseBeltItem(BeltSelSlot)
End Sub

Private Sub PicBelt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If BeltSelSlot < 1 Or BeltSelSlot > MaxBeltSlots Then
        Exit Sub
    End If
                          
    Select Case Button
    
        Dim Cuanto As Integer

        Case vbRightButton
        
            If Comerciando Then
            
                If UserMuerto Then 'Muerto
                    Exit Sub
                End If
                
                If Not MainTimer.Check(TimersIndex.BuySell) Then
                    Exit Sub
                End If

                If frmComerciar.Visible Then
                      
                    With frmComerciar
                        'If picBelt.Left + Left + X > .PicComercianteBelt.Left + .Left And _
                        'picBelt.Left + Left + X < .PicComercianteBelt.Left + .Left + .PicComercianteBelt.Width And _
                        'picBelt.Top + Top + Y > .PicComercianteBelt.Top + .Top And _
                        'picBelt.Top + Top + Y < .PicComercianteBelt.Top + .Top + .PicComercianteBelt.Height Then
                            
                            If LenB(.Cantidad.Text) < 1 Then
                                Exit Sub
                            End If
     
                            Cuanto = Val(Replace(.Cantidad.Text, ".", vbNullString))

                            If Cuanto > 0 And Cuanto <= MaxBeltObjs Then
                                If Cuanto > Belt(BeltSelSlot).Amount Then
                                    Cuanto = Belt(BeltSelSlot).Amount
                                End If
   
                                Call WriteCommerceSell(200 + BeltSelSlot, Cuanto)
                                Call Audio.mSound_PlayWav(SND_CLICK)
                            End If
                        'End If
                    End With

                Else
               
                    With frmBanco
                        'If X + picBelt.Left > .PicBancoBelt.Left And _
                        X + picBelt.Left < .PicBancoBelt.Left + .PicBancoBelt.Width And _
                        Y + picBelt.Top > .PicBancoBelt.Top And _
                        Y + picBelt.Top < .PicBancoBelt.Top + .PicBancoBelt.Height Then
                        
                            If LenB(.Cantidad.Text) < 1 Then
                                Exit Sub
                            End If

                            Cuanto = Val(Replace(.Cantidad.Text, ".", vbNullString))
                                                             
                            If Cuanto > 0 And Cuanto <= MaxBeltObjs Then
                                If Cuanto > Belt(BeltSelSlot).Amount Then
                                    Cuanto = Belt(BeltSelSlot).Amount
                                End If
                                
                                Call WriteBankDepositItem(200 + BeltSelSlot, Cuanto)
                                Call Audio.mSound_PlayWav(SND_CLICK)
                            End If
                        'End If
                    End With
                End If
                
            ElseIf Shift > 0 Then
                If BeltSelSlot > 0 And BeltSelSlot <= MaxBeltSlots Then
                
                    If Not MainTimer.Check(TimersIndex.Drop) Then
                        Exit Sub
                    End If
                
                    Call WriteDrop(200 + BeltSelSlot, 1)
                    Call Audio.mSound_PlayWav(SND_DROP)
                End If
            End If
        
        Case vbMiddleButton
        
            If X < 0 Or Y < 0 Or X > Width Or Y > Height Then
                Exit Sub
            End If
        
            If Not PicInvDragging Then
                If Not MainTimer.Check(TimersIndex.Drop) Then
                    Exit Sub
                End If
            
                If Belt(BeltSelSlot).Amount = 1 Or Shift > 0 Then
                    Call WriteDrop(200 + BeltSelSlot, Belt(BeltSelSlot).Amount)
                    Call Audio.mSound_PlayWav(SND_DROP)
                    
                Else
                    If Belt(BeltSelSlot).Amount > 1 Then
                        frmCantidadBelt.Show , Me
                    End If
                End If
        
                PicBelt.MousePointer = vbDefault
            End If

    End Select
    
    'If NpcInvSelSlot < 1 Then EXIT SUB
    
    'If DragType = MiBeltentario Then
        'If X < frmComerciar.Left + frmComerciar.PicComercianteBelt.Left Or _
        'X > frmComerciar.Left + frmComerciar.PicComercianteBelt.Left + frmComerciar.PicComercianteBelt.Height Then
        'DragType = None
        'MOUSEPOINTER = vbDefault
        'End If
    'End If

End Sub

Private Sub PicCompaInv_DblClick()
    
    If CompaSelSlot < 1 Then
        Exit Sub
    End If
    
    If UserMuerto Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
        Exit Sub
    End If

    Dim CompaName As String
    
    CompaName = Compa(CompaSelSlot).Nombre
 
    If SendTxt.Visible Then
    
        If InStr(CompaName, " ") > 0 Then
            CompaName = Replace(CompaName, " ", "+")
        End If
        
        If LenB(SendTxt.Text) > 0 Then
            
            If LenB(SendTxt.Text) = ":" & CompaName & " " Then
                If SendTxt.Text = ":" & CompaName & " " Then
                    SendTxt.Visible = False
                    SendTxt.Text = vbNullString
                    Exit Sub
                End If
            End If
            
            SendTxt.Text = ":" & CompaName & " " & SendTxt.Text
            
        Else
            SendTxt.Text = ":" & CompaName & " "
        End If
    Else
        SendTxt.Visible = True
    End If
    
    SendTxt.SetFocus
    'Call SendMessage (SendTxt.hWnd, &H7, ByVal 0&, ByVal 0&)
    SendTxt.SelStart = Len(SendTxt)
    
    Dim j As Integer
    
    For j = 1 To LastChar
        If Charlist(j).EsUser Then
            If LenB(Charlist(j).Nombre) = LenB(CompaName) Then
                If InStr(CompaName, "+") Then
                    CompaName = Replace(CompaName, "+", " ")
                End If
                
                If Charlist(j).Nombre = CompaName Then
                    SelectedCharIndex = j
                    Exit For
                End If
            End If
        End If
    Next j

End Sub

Private Sub picCompasScrollDown_Click()
    Dim sMax As Long
    Dim Pos As Long
    Dim Scroll_Info As SCROLLINFO
    
    Scroll_Info = GetScrollBarInfo(CompaRecTxt.hWnd)
    
    sMax = Scroll_Info.nMax - Scroll_Info.nPage + 1
    Pos = Scroll_Info.nTrackPos
    
    'If ret <> 0 Then
        If sMax = Pos Then
            MsgBox "It is the end."
        End If
    'End If

    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
End Sub

Private Sub picCompasScrollUp_Click()
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
End Sub

Private Sub PicInv_DblClick()

    If ButtonClicked = vbRightButton Then
        Exit Sub
    End If
    
    If frmCarp.Visible Then
        Exit Sub
    End If
    
    If frmHerrero.Visible Then
        Exit Sub
    End If
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
        Exit Sub
    End If
    
    If Comerciando Then
        Exit Sub
    End If
    
    If InvSelSlot < 1 Or InvSelSlot > MaxInvSlots Then
        Exit Sub
    End If
    
    If Inv(InvSelSlot).ObjIndex < 1 Then
        Exit Sub
    End If
    
    If UserMuerto Then
        'If Inv(Slot).ObjType <> otBarco Then
        '    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        '        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        '    End With
            Exit Sub
        'End If
    End If
    
    If MacroTrabajo.Enabled Then
        DesactivarMacroTrabajo
    End If
    
    If TrainingMacro.Enabled Then
        DesactivarMacroHechizos
    End If
    
    With Inv(InvSelSlot)
    
        'SI EL Item SOLO SE USA, MANDA USAR
        If .ObjType = otUseOnce Or _
            .ObjType = otGuita Or _
            .ObjType = otContenedor Or _
            .ObjType = otLlave Or _
            .ObjType = otPocion Or _
            .ObjType = otBebida Or _
            .ObjType = otPergamino Or _
            .ObjType = otBarco Or _
            .ObjType = otBotellaVacia Or _
            .ObjType = otBotellaLlena Then
                Call WriteUseItem(InvSelSlot)
                
            If .ObjType = otBebida Or .ObjType = otPocion Or .ObjType = otBotellaLlena Then
                Tomando = True
            End If
            
        ElseIf .ObjType = otMineral Then 'otMineral
            UsingSkill = 200 + InvSelSlot
            MousePointer = 2
        
        ElseIf .ObjIndex = 389 Or _
            .ObjIndex = 187 Or _
            .ObjIndex = 138 Or _
            .ObjIndex = 543 Or _
            .ObjIndex = 127 Or _
            .ObjIndex = 1005 Or _
            .ObjIndex = 198 Or _
            .ObjIndex = 15 Or _
            .ObjType = otCasco Or _
            .ObjType = otArmadura Or _
            .ObjType = otArma Or _
            .ObjType = otFlecha Or _
            .ObjType = otInstrumento Or _
            .ObjType = otEscudo Or _
            .ObjType = otCinturon Or _
            .ObjType = otAnillo Then

            Call Equipar
        End If
    End With
    
End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonClicked = Button
End Sub

Private Sub picBelt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonClicked = Button
End Sub

Private Sub PicInvCompa_Click()
    picInvCompa.Visible = False
    picListaCompa.Visible = True
    CompaRecTxt.Visible = False
    PicCompaInv.Visible = True
End Sub

Private Sub picListaCompa_Click()
    picListaCompa.Visible = False
    picInvCompa.Visible = True
    PicCompaInv.Visible = False
    CompaRecTxt.Visible = True
End Sub

Private Sub PicSpellInv_Click()

    If ButtonClicked = vbRightButton Then
        Exit Sub
    End If
    
    If frmCarp.Visible Then
        Exit Sub
    End If
        
    If frmHerrero.Visible Then
        Exit Sub
    End If
        
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
        Exit Sub
    End If
    
    If Meditando Then
        Exit Sub
    End If

    If Descansando Then
        Exit Sub
    End If
    
    If SpellSelSlot < 1 Then
        Exit Sub
    End If
    
    If SpellSelSlot > MaxSpellSlots Then
        Exit Sub
    End If
    
    If Spell(SpellSelSlot).Grh < 1 Then
        Exit Sub
    End If
    
    'If Not spell(SpellSelSlot).PuedeLanzar Then
    '    EXIT SUB
    'End If
    
    If UserMuerto Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
        Exit Sub
    End If
    
    If MacroTrabajo.Enabled Then
        DesactivarMacroTrabajo
    End If
    
    If TrainingMacro.Enabled Then
        DesactivarMacroHechizos
    End If

    Call Audio.mSound_PlayWav(SND_CLICK)
    
    'UsaMacro = True
    MousePointer = 2
    UsingSkill = Magia

End Sub

Private Sub PicSpellInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonClicked = Button
End Sub

Private Sub PicSpellInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not PicInvDragging Then
        If Button = vbRightButton Then
            If SpellSelSlot > 1 And SpellSelSlot <= MaxSpellSlots Then
                Call WriteSpellInfo(SpellSelSlot)
            End If
        End If
    End If
    
    PicSpellInv.MousePointer = vbDefault
End Sub

Private Sub LblPoblacion_Click()
    Call WriteOnline
End Sub

Private Sub Online_Click()
    Call WriteOnline
End Sub

Private Sub picConsoleScrollUp_Click()
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
End Sub

Private Sub picConsoleScrollDown_Click()
    Dim sMax As Long
    Dim Pos As Long
    Dim Scroll_Info As SCROLLINFO
    
    Scroll_Info = GetScrollBarInfo(RecTxt.hWnd)
    
    sMax = Scroll_Info.nMax - Scroll_Info.nPage + 1
    Pos = Scroll_Info.nTrackPos
    
    'If ret <> 0 Then
        If sMax = Pos Then
            MsgBox "It is the end."
        End If
    'End If

    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
End Sub

Private Sub picCompaScrollUp_Click()
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 0, ByVal 0&
End Sub

Private Sub picCompaScrollDown_Click()
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage CompaRecTxt.hWnd, &H115, 1, ByVal 0&
End Sub

Private Sub picHeadEqp_DblClick()
    Call DesEquipar(HeadEqp)
End Sub

Private Sub picBodyEqp_DblClick()
    Call DesEquipar(BodyEqp)
End Sub

Private Sub picLeftHandEqp_DblClick()
    Call DesEquipar(LeftHandEqp)
End Sub

Private Sub picRightHandEqp_DblClick()
    Call DesEquipar(RightHandEqp)
End Sub

Private Sub picBeltEqp_DblClick()
    Call DesEquipar(BeltEqp)
End Sub

Private Sub picRingEqp_DblClick()
    Call DesEquipar(RingEqp)
End Sub

Private Sub picShip_DblClick()
    Call DesEquipar(Ship)
End Sub

Private Sub picHeadEqp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If LenB(lblHeadEqp.Caption) > 0 Then
    '    lblHeadEqp.Visible = True
    'End If
End Sub

Private Sub picBodyEqp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If LenB(lblBodyEqp.Caption) > 0 Then
    '    lblBodyEqp.Visible = True
    'End If
End Sub

Private Sub picLeftHandEqp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If LenB(lblLeftHandEqp.Caption) > 0 Then
    '    lblLeftHandEqp.Visible = True
    'End If
End Sub

Private Sub picRightHandEqp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If LenB(lblRightHandEqp.Caption) > 0 Then
    '    lblRightHandEqp.Visible = True
    'End If
End Sub

Private Sub picBeltEqp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If LenB(lblBeltEqp.Caption) > 0 Then
    '    lblBeltEqp.Visible = True
    'End If
End Sub

Private Sub picRingEqp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If LenB(lblRingEqp.Caption) > 0 Then
    '    lblRingEqp.Visible = True
    'End If
End Sub

Private Sub picShip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If LenB(lblShip.Caption) > 0 Then
    '    lblShip.Visible = True
    'End If
End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub MacroTrabajo_Timer()
    If InvSelSlot = 0 Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    'If Not Application.IsAppActive() Then
    'DesactivarMacroTrabajo
    'EXIT SUB
    'End If
    
    If UsingSkill = Pescar Or UsingSkill = Talar Or UsingSkill = Minar Or (UsingSkill = Herreria And Not frmHerrero.Visible) Or UsingSkill > 200 Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If

    If frmCarp.Visible = False Then
        Call UseItemWithU
    End If
End Sub

Public Sub ActivarMacroTrabajo()
    MacroTrabajo.Interval = INT_MACRO_TRABAJO
    MacroTrabajo.Enabled = True
    Call ShowConsoleMsg("Macro Trabajo activado", 0, 100, 0, False, False, False)
End Sub

Public Sub DesactivarMacroTrabajo()
    MacroTrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call ShowConsoleMsg("Macro Trabajo desactivado", 100, 0, 0, False, False, False)
End Sub

Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If DragType = InventarioNpc Then
    'If NpcInvSelSlot > 0 Then
    'If IsNumeric(frmComerciar.Cantidad.Text) Then
    'If frmComerciar.Cantidad.Text > 0 Then
    'If UserGLD >= CalculateBuyPrice(NpcInv(NpcInvSelSlot).Valor, Val(frmComerciar.Cantidad.Text)) Then
    'Call WriteCommerceBuy(NpcInvSelSlot, frmComerciar.Cantidad.Text)
    'call Audio.mSound_PlayWav(SND_CLICK)
    'DragType = None
    'frmComerciar.PicComercianteInv.MOUSEPOINTER = vbDefault
    'Else
    'ShowConsoleMsg RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
    'End If
    'End If
    'End If
    'End If
    'End If
    
End Sub

Private Sub PicInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If InvSelSlot < 1 Or InvSelSlot > MaxInvSlots Then
        Exit Sub
    End If
                          
    Select Case Button
    
        Dim Cuanto As Integer

        Case vbRightButton
        
            If Comerciando Then
            
                If UserMuerto Then 'Muerto
                    Exit Sub
                End If
                
                If Not MainTimer.Check(TimersIndex.BuySell) Then
                    Exit Sub
                End If

                If frmComerciar.Visible Then
                      
                    With frmComerciar
                        'If picInv.Left + Left + X > .PicComercianteInv.Left + .Left And _
                        'picInv.Left + Left + X < .PicComercianteInv.Left + .Left + .PicComercianteInv.Width And _
                        'picInv.Top + Top + Y > .PicComercianteInv.Top + .Top And _
                        'picInv.Top + Top + Y < .PicComercianteInv.Top + .Top + .PicComercianteInv.Height Then
                            
                            If LenB(.Cantidad.Text) < 1 Then
                                Exit Sub
                            End If
     
                            Cuanto = Val(Replace(.Cantidad.Text, ".", vbNullString))

                            If Cuanto > 0 And Cuanto <= MaxInvObjs Then
                                If Cuanto > Inv(InvSelSlot).Amount Then
                                    Cuanto = Inv(InvSelSlot).Amount
                                End If
   
                                Call WriteCommerceSell(InvSelSlot, Cuanto)
                                Call Audio.mSound_PlayWav(SND_CLICK)
                            End If
                        'End If
                    End With

                Else
               
                    With frmBanco
                        'If X + picInv.Left > .PicBancoInv.Left And _
                        X + picInv.Left < .PicBancoInv.Left + .PicBancoInv.Width And _
                        Y + picInv.Top > .PicBancoInv.Top And _
                        Y + picInv.Top < .PicBancoInv.Top + .PicBancoInv.Height Then
                        
                            If LenB(.Cantidad.Text) < 1 Then
                                Exit Sub
                            End If

                            Cuanto = Val(Replace(.Cantidad.Text, ".", vbNullString))
                                                             
                            If Cuanto > 0 And Cuanto <= MaxInvObjs Then
                                If Cuanto > Inv(InvSelSlot).Amount Then
                                    Cuanto = Inv(InvSelSlot).Amount
                                End If
                                
                                Call WriteBankDepositItem(InvSelSlot, Cuanto)
                                Call Audio.mSound_PlayWav(SND_CLICK)
                            End If
                        'End If
                    End With
                End If
                
            ElseIf Shift > 0 Then
                If InvSelSlot > 0 And InvSelSlot <= MaxInvSlots Then
                
                    If Not MainTimer.Check(TimersIndex.Drop) Then
                        Exit Sub
                    End If
                
                    Call WriteDrop(InvSelSlot, 1)
                    Call Audio.mSound_PlayWav(SND_DROP)
                End If
            End If
        
        Case vbMiddleButton
        
            If X < 0 Or Y < 0 Or X > Width Or Y > Height Then
                Exit Sub
            End If
        
            If Not PicInvDragging Then
                If Not MainTimer.Check(TimersIndex.Drop) Then
                    Exit Sub
                End If
            
                If Inv(InvSelSlot).Amount = 1 Or Shift > 0 Then
                    Call WriteDrop(InvSelSlot, Inv(InvSelSlot).Amount)
                    Call Audio.mSound_PlayWav(SND_DROP)
                    
                Else
                    If Inv(InvSelSlot).Amount > 1 Then
                        frmCantidad.Show , Me
                    End If
                End If
        
                PicInv.MousePointer = vbDefault
            End If

    End Select
    
    'If NpcInvSelSlot < 1 Then EXIT SUB
    
    'If DragType = MiInventario Then
        'If X < frmComerciar.Left + frmComerciar.PicComercianteInv.Left Or _
        'X > frmComerciar.Left + frmComerciar.PicComercianteInv.Left + frmComerciar.PicComercianteInv.Height Then
        'DragType = None
        'MOUSEPOINTER = vbDefault
        'End If
    'End If
End Sub

Private Sub picRightHandEqp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        Exit Sub
    End If
    
    If RightHandEqp.ObjIndex = 0 Then
        Exit Sub
    End If
    
    If Not RightHandEqp.Proyectil Then
        Exit Sub
    End If

    If UserMuerto Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
        Exit Sub
    End If

    If MacroTrabajo.Enabled Then
        DesactivarMacroTrabajo
    End If
   
    If TrainingMacro.Enabled Then
        DesactivarMacroHechizos
    End If
    
    If Meditando Then
        Exit Sub
    End If
    
    If Descansando Then
        Exit Sub
    End If
    
    UsingSkill = Proyectiles
    MousePointer = 2
End Sub

Private Sub RecTxt_GotFocus()
    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        PicSpellInv.SetFocus
    End If
End Sub

Private Sub RecTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim txt As String
    
    txt = RichWordOver(RecTxt, X, Y)
    
    If LenB(txt) > 0 Then
        RecTxt.MousePointer = rtfArrowQuestion
        RecTxt.ToolTipText = "Hablar con " & txt
    Else
        If RecTxt.MousePointer <> 1 Then
            RecTxt.MousePointer = 1
            RecTxt.ToolTipText = vbNullString
        End If
    End If
    
    MouseOverRecTxt = True
    
End Sub

Private Sub RecTxt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim CompaName As String
    
    CompaName = RichWordOver(RecTxt, X, Y)
    
    If LenB(CompaName) < 3 Then
        Exit Sub
    End If
    
    If Button = vbRightButton Then
        If Charlist(UserCharIndex).Priv > 1 Then
            MsgBox CompaName
            Call WriteGoToChar(CompaName)
            Exit Sub
        End If
    End If

    Dim j As Integer

    If SendTxt.Visible Then

        If SendTxt.Text = ":" & Replace(CompaName, " ", "+") & " " Then
            SendTxt.Visible = False
            SendTxt.Text = vbNullString
            
            If SelectedCharIndex > 0 Then
                If Charlist(SelectedCharIndex).EsUser Then
                    SelectedCharIndex = 0
                End If
            End If
        
        Else
            SendTxt.Text = ":" & Replace(CompaName, " ", "+") & " " & SendTxt.Text
            SendTxt.SetFocus
            SendTxt.SelStart = Len(SendTxt)
                            
            For j = 1 To LastChar
                If Charlist(j).EsUser Then
                    If LenB(Charlist(j).Nombre) = LenB(CompaName) Then
                        If Charlist(j).Nombre = CompaName Then
                            SelectedCharIndex = j
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If
            
    Else
        SendTxt.Text = ":" & Replace(CompaName, " ", "+") & " "
    
        SendTxt.Visible = True
        SendTxt.SetFocus
        SendTxt.SelStart = Len(SendTxt)
                     
        For j = 1 To LastChar
            If Charlist(j).EsUser Then
                If LenB(Charlist(j).Nombre) = LenB(CompaName) Then
                    If Charlist(j).Nombre = CompaName Then
                        SelectedCharIndex = j
                        Exit For
                    End If
                End If
            End If
        Next j
    End If
End Sub

Private Sub SendTxt_GotFocus()
    'SendTxtBorder.Visible = True
End Sub

Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(Trim$(SendTxt.Text)) < 1 Then
            Call WriteDeleteChat
            Call Dialogos.RemoveDialog(UserCharIndex)
        Else
            Call ParseUserCommand(Trim$(SendTxt.Text))
        End If
        
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.Visible = False
        
        If SelectedCharIndex > 0 Then
            If Charlist(SelectedCharIndex).EsUser Then
                SelectedCharIndex = 0
            End If
        End If
        
        'RecTxt.SetFocus
        
    ElseIf KeyCode = vbKeyUp And Shift > 0 Then
        If LenB(SendTxt.Text) < 1 And LenB(LastParsedString) > 0 Then
            SendTxt.Text = LastParsedString
            SendTxt.SelStart = Len(LastParsedString)
        End If
    
    ElseIf KeyCode = vbKeyBack Then
        If Left$(SendTxt.Text, 1) = ":" Then
            Dim Epa() As String
            Epa = Split(SendTxt.Text, " ", 2)

            If LenB(SendTxt.Text) = LenB(Epa(0)) + 2 Then
                If SendTxt.Text = Epa(0) & " " Then
                    SendTxt.Text = vbNullString
                    
                    If SelectedCharIndex > 0 Then
                        If Charlist(SelectedCharIndex).EsUser Then
                            SelectedCharIndex = 0
                        End If
                    End If
                End If
            End If
        End If
    
    ElseIf KeyCode = vbKeyEscape Then
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub

Private Sub UseItemWithU()

    If UserMuerto Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
        Exit Sub
    End If

    'If macrotrabajo.Enabled Then
    'DesactivarMacroTrabajo
    'End If
   
    If TrainingMacro.Enabled Then
        DesactivarMacroHechizos
    End If
        
    If InvSelSlot < 1 Or InvSelSlot > MaxInvSlots Then
        Exit Sub
    End If
    
    If Comerciando Then
        Exit Sub
    End If
    
    With Inv(InvSelSlot)
    
        'SI EL Item SOLO SE USA, MANDA USAR
        If .ObjType = otUseOnce Or _
            .ObjType = otGuita Or _
            .ObjType = otContenedor Or _
            .ObjType = otLlave Or _
            .ObjType = otPocion Or _
            .ObjType = otBebida Or _
            .ObjType = otPergamino Or _
            .ObjType = otBarco Or _
            .ObjType = otBotellaVacia Or _
            .ObjType = otBotellaLlena Then
            
            Call WriteUseItem(InvSelSlot)
            
            If .ObjType = otBebida Or .ObjType = otPocion Or .ObjType = otBotellaLlena Then
                Tomando = True
            End If
            
        ElseIf .ObjType = otMineral Then 'otMineral
            UsingSkill = 200 + InvSelSlot
            MousePointer = 2
        
        ElseIf .ObjIndex = 389 Or _
            .ObjIndex = 187 Or _
            .ObjIndex = 138 Or _
            .ObjIndex = 543 Or _
            .ObjIndex = 127 Or _
            .ObjIndex = 1005 Or _
            .ObjIndex = 198 Or _
            .ObjIndex = 15 Or _
            .ObjType = otCasco Or _
            .ObjType = otArmadura Or _
            .ObjType = otArma Or _
            .ObjType = otFlecha Or _
            .ObjType = otInstrumento Or _
            .ObjType = otEscudo Or _
            .ObjType = otCinturon Or _
            .ObjType = otAnillo Then

            Call Equipar
        End If
        
    End With

End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If Left$(SendTxt.Text, 1) = ":" Then
        Dim Es As Boolean
        Dim CompaName As String
        Dim i As Byte

        For i = 1 To MaxCompaSlots
            If Compa(i).Online Then
                CompaName = Compa(i).Nombre
                
                If LenB(Left$(SendTxt.Text, Len(CompaName) + 2)) = LenB(":" & CompaName & " ") Then
                    If Left$(SendTxt.Text, Len(CompaName) + 2) = ":" & Replace(CompaName, " ", "+") & " " Then
                        Es = True
                        Exit For
                    End If
                End If
            End If
        Next i
                
        If SelectedCharIndex > 0 Then
            If Not Es Then
                If Charlist(SelectedCharIndex).EsUser Then
                    SelectedCharIndex = 0
                End If
            End If
        
        Else
            Dim j As Integer
        
            If Es Then
                For j = 1 To LastChar
                    If Charlist(j).EsUser Then
                        If LenB(Charlist(j).Nombre) = LenB(CompaName) Then
                            If Charlist(j).Nombre = CompaName Then
                                SelectedCharIndex = j
                                Exit For
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    End If
    
End Sub

Private Sub SendTxt_LostFocus()
    'SendTxtBorder.Visible = False
    
    If LenB(SendTxt.Text) > 0 Then
        If NroCompas > 0 Then
            Dim i As Byte

            For i = 1 To MaxCompaSlots
        
                If Compa(i).Online Then
                
                    Dim Name As String
                    
                    Name = Compa(i).Nombre
                                                    
                    If InStr(Name, " ") > 0 Then
                        Name = Replace(Name, " ", "+")
                    End If
                            
                    If LenB(Left$(SendTxt.Text, Len(Name) + 2)) = LenB(":" & Name & " ") Then
                        If Left$(SendTxt.Text, Len(Name) + 2) = ":" & Name & " " Then
                            SendTxt.Visible = False
                            SendTxt.Text = vbNullString
                            
                            If SelectedCharIndex > 0 Then
                                If Charlist(SelectedCharIndex).EsUser Then
                                    SelectedCharIndex = 0
                                End If
                            End If
                            
                            Exit For
                        End If
                    End If
                    
                End If
            Next i
        End If
    End If
End Sub

Private Sub Socket1_Disconnect()

    If frmCrearPersonaje.Visible Then
        frmCrearPersonaje.MousePointer = 0
    End If
    
    With frmConnect
        .PasswordTxt.Visible = True
        .NameTxt.Visible = True
        .MousePointer = 0
            
        If .PasswordTxtBorder.BorderColor = &H80& Or .NameTxtBorder.BorderColor = &H80& Then
            Exit Sub
        End If
            
    End With
    
    Dim i As Integer
    
    'Hide main form
    Visible = False
    frmConnect.Visible = True
    
    UserLogged = False
    
    LastParsedString = vbNullString
    
    'Stop audio
    Call Audio.mSound_StopWav(0)
    IsPlaying = PlayLoop.plNone
    
    Call Audio.mMusic_StopMid
    
    Call Audio.MusicMP3Stop
        
    SendTxt.Visible = False
    SendTxt.Text = vbNullString
        
    Descansando = False
    UserParalizado = False
    Pausa = False
    UserCiego = False
    Meditando = False
    UserNavegando = False
    bRain = False
    bFogata = False
    SkillPoints = 0
    
    ImgExp.Width = 0
    ImgHP.Width = 0
    ImgMana.Width = 0
    ImgSta.Width = 0
    
    UserMap = 0
    
    UserCharIndex = 0
    
    SelectedCharIndex = 0
    
    InvSelSlot = 0
    BeltSelSlot = 0
    SpellSelSlot = 0
    CompaSelSlot = 0
    NpcInvSelSlot = 0
    
    TempSlot = 0
    
    Label.Caption = vbNullString
    
    picHeadEqp.Picture = picHeadEqp.Picture
    picBodyEqp.Picture = picBodyEqp.Picture
    picLeftHandEqp.Picture = picLeftHandEqp.Picture
    picRightHandEqp.Picture = picRightHandEqp.Picture
    picBeltEqp.Picture = picBeltEqp.Picture
    picRingEqp.Picture = picRingEqp.Picture
    'picShip.Picture = picShip.Picture
    
    PicSpellInv.Visible = False
    PicInv.Visible = True
    
    NroItems = 0

    Set PicInv.Picture = PicInv.Picture
        
    For i = 1 To MaxInvSlots 'NroItems
        If Inv(i).ObjIndex > 0 Then
            Call Inventario.UnSetSlot(i, False)
        End If
    Next i
        
    NroBeltItems = 0
    
    Set PicBelt.Picture = PicBelt.Picture
    
    For i = 1 To MaxBeltSlots 'NroBeltItems
        If Belt(i).ObjIndex > 0 Then
            Call Cinturon.UnSetSlot(i, False)
        End If
    Next i
    
    NroSpells = 0
    
    Set PicSpellInv.Picture = PicSpellInv.Picture
    
    For i = 1 To MaxSpellSlots
        If Spell(i).Grh > 0 Then
            Call Hechizos.UnSetSlot(i, False)
        End If
    Next i
    
    NroCompas = 0
            
    For i = 1 To MaxCompaSlots
        If LenB(Compa(i).Nombre) > 0 Then
            Call UnSetCompaSlot(i)
        End If
    Next i
        
    For i = 1 To MaxPlataformSlots
        Plataforma(i) = 0
    Next i
        
    CompaRecTxt.Text = vbNullString
        
    RecTxt.Text = vbNullString
    
    Call Dialogos.RemoveAllDialogs
    
    'Reset some char variables...
    For i = 1 To LastChar
        Charlist(i).Invisible = False
        Charlist(i).Paralizado = False
    Next i
        
    'Unload all forms except me and frmConnect
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name <> Name And frm.Name <> frmConnect.Name Then
            Unload frm
        End If
    Next
    
    UserPassword = vbNullString
            
    Dim eName As String
    Dim ePass As String
        
    'Get the username/password
    eName = (GetVar(DataPath & "Game.ini", "INIT", "Name"))
    ePass = (GetVar(DataPath & "Game.ini", "INIT", "Pass"))
    
    With frmConnect
        
        If LenB(UserName) > 0 Then
            .NameTxt.Text = UserName
  
            If UserName = eName And LenB(ePass) > 0 Then
                .PasswordTxt.Text = ePass
                .SavePassImg.Visible = True
                Call SendMessage(.NameTxt.hWnd, &H7, ByVal 0&, ByVal 0&)
            Else
                UserPassword = vbNullString
                
                .PasswordTxt.SelStart = 0
                .PasswordTxt.SelLength = Len(.PasswordTxt.Text)
                Call SendMessage(.PasswordTxt.hWnd, &H7, ByVal 0&, ByVal 0&)
            End If
        Else
            Call SendMessage(.NameTxt.hWnd, &H7, ByVal 0&, ByVal 0)
        End If
                
        Visible = False
        
        'Show connection form
        frmConnect.Visible = True
                
        'Call SetWindowPos(frmConnect.hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
                
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    
        If LenB(UserName) > 0 Then
            Call SetWindowPos(frmConnect.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
        End If
    
    End With

End Sub

Private Sub Timer1_Timer()
    If SendTxt.Visible Then
        If LenB(SendTxt.Text) > 0 Then
            If Not Screen.ActiveControl Is Nothing Then
                If Screen.ActiveControl = SendTxt Then
                    Dim a As Byte
                    a = SendTxt.SelStart
                    
                    SendTxt.SelStart = 0
                
                    SendTxt.SelLength = Len(SendTxt.Text)
                        
                    SendTxt.SelColor = vbWhite
                    
                    SendTxt.SelStart = a
                End If
            End If
        End If
    End If
End Sub

Private Sub TrainingMacro_Timer()

    If Not PicSpellInv.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    'Macros are disabled if focus is not on Argentum!
    'If Not Application.IsAppActive() Then
    'DesactivarMacroHechizos
    'EXIT SUB
    'End If
    
    If PicSpellInv.Visible = False Or SpellSelSlot < 1 Or UserMinMan < 1 Then
        DesactivarMacroHechizos
        MousePointer = 0
        Exit Sub
    End If

    If MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
        If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then
            Exit Sub
        End If
        
        Call WriteCastSpell(SpellSelSlot, tX, tY)
    End If
End Sub

Private Sub Form_DblClick()
    'If Not frmComerciar.Visible And Not frmComerciarUsu.Visible And _
    Not frmBanco.Visible And Not frmSkills.Visible And _
    Not frmMSG.Visible And Not frmEntrenador.Visible And _
    Not frmEstadisticas.Visible And Not frmCantidad.Visible And Not frmCantidadGld.Visible Then
    'Call WriteDoubleClick(tx, tY)
    'End If
End Sub

Private Sub Form_Load()

On Error Resume Next

    Left = 0
    Top = 0
    ImgExp.Width = 0
    ImgHP.Width = 0
    ImgMana.Width = 0
    ImgSta.Width = 0

    Picture = LoadPicture(GrhPath & "Principal.jpg")

    'RECTXT TRANSPARENTE
    Call Make_Transparent_Richtext(RecTxt.hWnd)
    
    'SENDTXT TRANSPARENTE
    Call Make_Transparent_Richtext(SendTxt.hWnd)
    
    'COMPARECTXT TRANSPARENTE
    Call Make_Transparent_Richtext(CompaRecTxt.hWnd)
    
    Call EnableURLDetect(RecTxt.hWnd, hWnd)
    
    Call Round_Picture(Minimap, Minimap.Width)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
On Error Resume Next
    Exit Sub
    MouseOverRecTxt = False
    
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top
        
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewShp.Width Then
        MouseX = MainViewShp.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If

    If InGameArea() Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    Else
        'lblHeadEqp.Visible = False
        'lblBodyEqp.Visible = False
        'lblLeftHandEqp.Visible = False
        'lblRightHandEqp.Visible = False
        'lblBeltEqpEqp.Visible = False
        'lblRingEqpEqp.Visible = False
        'lblShip.Visible = False

        MouseTileX = 0
        MouseTileY = 0
    End If
    
    'If MouseX < Minimap.Left Or MouseX > Minimap.Left + Minimap.Width Then
    'If MouseY < Minimap.Top Or clicY > Minimap.Top + Minimap.Height Then
    'Minimap.Visible = True
    'End If
    'End If

    If LenB(Label.Caption) > 0 Then
        Label.Caption = vbNullString
    End If
    
    Dim TempSlot2 As Byte
    
    If TempSlot > 0 Then
        TempSlot2 = TempSlot
        TempSlot = 0
        
        If Belt(TempSlot2).ObjIndex > 0 Then
            If TempSlot2 <> BeltSelSlot Then
                Call Cinturon.DrawBeltSlot(TempSlot2)
            End If
        End If
        
        If PicInv.Visible Then
            If Inv(TempSlot2).ObjIndex > 0 Then
                Call Inventario.DrawSlot(TempSlot2)
            End If
        Else
            If Spell(TempSlot2).Grh > 0 Then
                Call Hechizos.DrawSpellSlot(TempSlot2)
            End If
        End If
    
    ElseIf BeltTempSlot > 0 Then
        TempSlot2 = BeltTempSlot
        BeltTempSlot = 0
        
        If Belt(TempSlot2).ObjIndex > 0 Then
            Call Cinturon.DrawBeltSlot(TempSlot2)
        End If
    End If
    
    If UserPasarNivel > 0 Then
        ExpLbl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
    End If
End Sub

Private Sub imgDropGold_Click()
    If UserGld > 0 Then
        frmCantidadGld.Show , Me
    End If
End Sub

Private Sub imgAsignarSkill_Click()
    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills.puntos.Caption = SkillPoints
    frmSkills.Show , Me
End Sub

Private Sub lblInventario_Click()
    If PicInv.Visible Then
        Call Audio.mSound_PlayWav(SND_CLICK)
    
        imgInv.Visible = False
        PicSpellInv.Visible = False
        
        PicInv.Visible = True
    End If
End Sub

Private Sub lblHechizos_Click()
    If Not imgInv.Visible Then
        Call Audio.mSound_PlayWav(SND_CLICK)
    
        PicInv.Visible = False
        
        imgInv.Visible = True
        PicSpellInv.Visible = True
    End If
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba Errores al salir y volver a entrar
    If Not modApplication.IsAppActive() Then
        Exit Sub
    End If
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    'ElseIf Not frmComerciar.Visible And Not frmComerciarUsu.Visible And _
    Not frmBanco.Visible And Not frmSkills.Visible And _
    Not frmMSG.Visible And Not frmEntrenador.Visible And _
    Not frmEstadisticas.Visible And Not frmCantidad.Visible And Not frmCantidadGld.Visible Then
    End If

    If MouseOverRecTxt Then
        Call SetVerticalScrollPos(RecTxt, Vertical_Pos)
    Else
        SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
    End If
    
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 1, ByVal 0&
    SendMessage RecTxt.hWnd, &H115, 0, ByVal 0&
End Sub

Private Sub SendTxt_Change()
    
    If LenB(SendTxt.Text) < 1 Then
        Exit Sub
    End If
    
    Dim Sel As Byte
    
    Sel = SendTxt.SelStart
    
    Dim Text As String
    
    Text = SendTxt.Text
    
    Text = Replace(Text, "  ", " ")
    Text = Replace(Text, " .", ".")
    
    If Left$(Text, 1) = ":" Then
        If InStr(Text, " ") > 0 Then
            If Left$(mid$(Text, 2), InStr(Text, " ") - 2) <> StrConv(Left$(mid$(Text, 2), InStr(Text, " ") - 2), vbProperCase) Then
                Text = ":" & StrConv(Left$(mid$(Text, 2), InStr(Text, " ") - 2), vbProperCase)
            End If
        Else
            If mid$(Text, 2) <> StrConv(mid$(Text, 2), vbProperCase) Then
                Text = ":" & StrConv(mid$(Text, 2), vbProperCase)
            End If
        End If
        
    ElseIf Left$(Text, 1) = "+" Or Left$(Text, 1) = "-" Then
        'If InStr(text, " ") > 0 Then
        '    text = Replace(text, " ", "+")
        'End If
        
        If mid$(Text, 2) <> StrConv(mid$(Text, 2), vbProperCase) Then
            Text = Left$(Text, 1) & StrConv(mid$(Text, 2), vbProperCase)
        End If
        
    Else
        If Len(Text) > 3 Then
            If Left$(Text, 1) <> UCase$(Left$(Text, 1)) Then
                Text = UCase$(Left$(Text, 1)) & mid$(Text, 2)
            End If
        
            If Left$(Text, 3) = UCase$(Left$(Text, 3)) Then
                Text = UCase$(Left$(Text, 1)) & mid$(Text, 2)
            End If
            
            If mid$(Text, Len(Text) - 1, 1) = "." Then
                If Right$(Text, 1) <> UCase$(Right$(Text, 1)) Then
                    Text = Left$(Text, Len(Text) - 1) & " " & UCase$(Right$(Text, 1))
                    Sel = Sel + 1
                End If
            End If
            
            If mid$(Text, Len(Text) - 2, 2) = ". " Then
                If Right$(Text, 1) <> UCase$(Right$(Text, 1)) Then
                    Text = Left$(Text, Len(Text) - 1) & UCase$(Right$(Text, 1))
                    Sel = Sel + 1
                End If
            End If
        End If
    End If
    
    SendTxt.Text = Text
    
    SendTxt.SelStart = 0

    SendTxt.SelLength = Len(Text)
        
    SendTxt.SelColor = vbWhite
    
    SendTxt.SelStart = Sel
    
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If KeyAscii < vbKeySpace Or KeyAscii > 250 Then
        KeyAscii = 0
    End If
End Sub

'Socket1
Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
       
    Select Case EstadoLogin
        Case Creado, Normal, Recuperando
            Call Login
        
        Case BuscandoNombre
            Call WriteRequestRandomName
            
            DoEvents
            Call FlushBuffer
        
        Case Creando
            frmCrearPersonaje.Show vbModal
    End Select
    
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    'Handle socket Errors
        
    If ErrorCode = 24053 Or ErrorCode = 24036 Then
        Exit Sub
    End If

    If ErrorCode = 24061 Then
        MsgBox "Abraxas se encuentra fuera de línea. Intentá nuevamente más tarde."
    ElseIf ErrorCode = 25004 Then
        MsgBox "Abraxas se encuentra fuera de línea. Intentá nuevamente más tarde."
        Call CloseClient
    ElseIf ErrorCode = 24060 Then
        MsgBox "Tiempo de espera agotado. Verificá tu conexión a internet e intentá nuevamente."
    Else
        Call MsgBox(ErrorCode & ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    End If
    
    frmConnect.MousePointer = 0
    Response = 0

    If Socket1.Connected Then
        Socket1.Disconnect
    End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)

On Error Resume Next

    Dim RD As String
    Dim data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)
    
    If LenB(RD) < 1 Then
        Exit Sub
    End If
    
    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Function InGameArea() As Boolean
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + MainViewShp.Width Then
        Exit Function
    End If
    
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + MainViewShp.Height Then
        Exit Function
    End If
    
    InGameArea = True
End Function

Private Sub Exp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserPasarNivel > 0 Then
        ExpLbl.Caption = UserExp & " / " & UserPasarNivel
    End If
End Sub
Private Sub Minimap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call frmMapa.Show
    Else
        If X < 10 Or X > 91 Or Y < 8 Or Y > 93 Then
            Exit Sub
        End If
        
        If Charlist(UserCharIndex).Priv > 1 Then
            Call WriteWarpChar("YO", UserMap, CByte(X), CByte(Y))
        End If
    End If
End Sub

'PRIVATE SUB Minimap_MouseMove(Button As Integer, Shift As Integer, x as Single, y as Single)
'If Charlist(UserCharIndex).Priv < 2 Then
'Minimap.Visible = False
'End If
'END SUB

'Return the word the mouse is over.

Public Function RichWordOver(rch As RichTextBox, X As Single, Y As Single) As String
    
On Error GoTo Error

    Dim pt As POINTAPI
    Dim Pos As Integer
    Dim start_pos As Integer
    Dim end_pos As Integer
    Dim ch As String
    Dim txt As String
    Dim txtlen As Integer

    'Convert the position to pixels.
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    'Get the Char number
    Pos = SendMessage(rch.hWnd, &HD7, 0&, pt)
    
    If Pos <= 0 Then
        Exit Function
    End If
    
    'Find the start of the word.
    txt = rch.Text
    
    If rch = RecTxt Then
        Exit Function
        For start_pos = Pos To 1 Step -1
        
            ch = mid$(rch.Text, start_pos + 1, 1)
            'Allow digits, letters, and underscores.
            If ch = UCase$(ch) Then
                Exit For
            End If
            
        Next start_pos
    
    Else
    
        For start_pos = Pos To 1 Step -1
        
            ch = mid$(rch.Text, start_pos, 1)
            'Allow digits, letters, and underscores.
            If Not ((ch >= "0" And ch <= "9") Or _
                (ch >= "a" And ch <= "z") Or _
                (ch >= "A" And ch <= "Z")) Then
                Exit For
            End If
            
        Next start_pos
    
    End If
    
    start_pos = start_pos + 1

    'Find the end of the word.
    txtlen = Len(txt)
    
    If rch = RecTxt Then
        For end_pos = Pos To txtlen
        
            ch = mid$(txt, end_pos, 1)
            
            'Allow digits, letters, and underscores.
            If ch = ":" Or end_pos > 15 + start_pos Then
                Exit For
            End If
            
        Next end_pos

    Else
        For end_pos = Pos To txtlen
        
            ch = mid$(txt, end_pos, 1)
            
            'Allow digits, letters, and underscores.
            If Not ((ch >= "0" And ch <= "9") Or _
                (ch >= "a" And ch <= "z") Or _
                (ch >= "A" And ch <= "Z")) Then
                Exit For
            End If
            
        Next end_pos
    
    End If
    
    end_pos = end_pos - 1

    If start_pos <= end_pos Then
        RichWordOver = mid$(txt, start_pos, end_pos - start_pos + 1)
        
        If rch.Tag = RecTxt.Tag Then
            If LenB(RichWordOver) = LenB(UserName) Then
                If RichWordOver = UserName Then
                    RichWordOver = vbNullString
                    Exit Function
                End If
            End If
            
            If EsCompaniero(RichWordOver) > 0 Then
                'RecTxt.SelStart = start_pos - 1
                'RecTxt.SelLength = end_pos
                'RecTxt.SelColor = vbRed
            ElseIf mid$(txt, end_pos + 1, 1) = ":" Then
                'RecTxt.SelStart = start_pos - 1
                'RecTxt.SelLength = end_pos
                'RecTxt.SelColor = vbRed
            Else
                Dim a As Boolean
                Dim j As Integer
                
                For j = 1 To LastChar
                    If Charlist(j).EsUser Then
                        If LenB(Charlist(j).Nombre) = LenB(RichWordOver) Then
                            If Charlist(j).Nombre = RichWordOver Then
                                SelectedCharIndex = j
                                a = True
                                Exit For
                            End If
                        End If
                    End If
                Next j
            
                If Not a Then
                    RichWordOver = vbNullString
                End If
            End If
        'Else
            'CompaRecTxt.SelStart = start_pos - 1
            'CompaRecTxt.SelLength = end_pos
            'CompaRecTxt.SelColor = vbRed
        End If
        
    End If
    
    Debug.Print RichWordOver

Error:
End Function
