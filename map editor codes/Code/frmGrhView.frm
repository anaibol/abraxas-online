VERSION 5.00
Begin VB.Form frmMode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   616
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Particulas"
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   36
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Mapa"
      Height          =   375
      Index           =   9
      Left            =   1680
      TabIndex        =   77
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Automatizadores"
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   118
      Top             =   8760
      Width           =   1575
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Opciones"
      Height          =   735
      Index           =   8
      Left            =   3480
      TabIndex        =   76
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Traslados"
      Height          =   375
      Index           =   6
      Left            =   2640
      TabIndex        =   46
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Triggers"
      Height          =   375
      Index           =   7
      Left            =   1800
      TabIndex        =   47
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Luces"
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   35
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Bloqueos"
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   38
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Objetos"
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   37
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "NPCs"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   34
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Superficies"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame fGrhView 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame fSups 
         Caption         =   "Insertar Superficies"
         Height          =   5415
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4095
         Begin VB.CheckBox chPutSup 
            Caption         =   "Insertar/Quitar"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox lstSuperfices 
            Height          =   1620
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   3855
         End
         Begin VB.PictureBox picSups 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00000080&
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
            Height          =   1920
            Left            =   120
            ScaleHeight     =   128
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   256
            TabIndex        =   10
            Top             =   3000
            Width           =   3840
         End
         Begin VB.CheckBox chOrdenar 
            Caption         =   "ordernar"
            Height          =   195
            Left            =   1920
            TabIndex        =   9
            Top             =   5040
            Width           =   975
         End
         Begin VB.CheckBox ckView 
            Caption         =   "Visualizar"
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txSearchSurface 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   2610
            Width           =   3855
         End
         Begin VB.CheckBox AutoCompletarSuperficie 
            Caption         =   "Auto"
            Height          =   195
            Left            =   3000
            TabIndex        =   6
            Top             =   5040
            Width           =   855
         End
         Begin VB.ListBox lstCapa 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            ItemData        =   "frmGrhView.frx":0000
            Left            =   3240
            List            =   "frmGrhView.frx":0010
            TabIndex        =   5
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblStatSup 
            Caption         =   "Capa:"
            Height          =   255
            Left            =   2760
            TabIndex        =   22
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Layer:"
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alto : "
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   5040
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho : "
            Height          =   195
            Left            =   720
            TabIndex        =   14
            Top             =   5040
            Width           =   600
         End
         Begin VB.Label lblS 
            Alignment       =   2  'Center
            Caption         =   "Buscar"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2400
            Width           =   3855
         End
      End
      Begin VB.CommandButton cmdSupInsertSelec 
         Caption         =   "Insertar Superficie en espacio seleccionado"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   6240
         Width           =   3855
      End
      Begin VB.CommandButton cmdSupInsertMap 
         Caption         =   "Insertar esta superficie en todo el mapa"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   7200
         Width           =   3855
      End
      Begin VB.CommandButton cmdSupSacSelec 
         Caption         =   "Sacar Superficie en espacio seleccionado"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   6720
         Width           =   3855
      End
   End
   Begin VB.Frame fNPCs 
      Height          =   7695
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   4350
      Begin VB.CommandButton cmdInsertNpcMap 
         Caption         =   "Insertar al azar"
         Height          =   495
         Left            =   2640
         TabIndex        =   113
         Top             =   2760
         Width           =   1575
      End
      Begin VB.PictureBox picRenderNpc 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00000080&
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
         Height          =   3000
         Left            =   120
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   21
         Top             =   240
         Width           =   2400
      End
      Begin VB.ListBox lstNPC 
         Height          =   1815
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chPutNpc 
         Caption         =   "Insertar/Quitar"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Height          =   1695
         Left            =   1800
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fLuces 
      Height          =   7695
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox chPutLight 
         Caption         =   "Insertar/Quitar"
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRango 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   30
         Text            =   "5"
         Top             =   840
         Width           =   2895
      End
      Begin VB.Frame t 
         Caption         =   "Color"
         Height          =   975
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   2895
         Begin VB.TextBox txtAlpha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   28
            Text            =   "255"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtRed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   27
            Text            =   "255"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtGreen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   26
            Text            =   "255"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtBlue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   25
            Text            =   "255"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Alpha        Rojo        Verde        Azul"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Rango"
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   600
         Width           =   2880
      End
   End
   Begin VB.Frame fTraslado 
      Height          =   7695
      Left            =   120
      TabIndex        =   66
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox chPutTraslado 
         BackColor       =   &H80000004&
         Caption         =   "Insertar/Quitar"
         Height          =   240
         Left            =   120
         TabIndex        =   71
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtMap 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         TabIndex        =   70
         Text            =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtX 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         TabIndex        =   69
         Text            =   "1"
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtY 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         TabIndex        =   68
         Text            =   "1"
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdShowUnion 
         Caption         =   "Union de mapas ady"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Mapa a.."
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "X a.."
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Y a.."
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame fOpc 
      Height          =   7695
      Left            =   120
      TabIndex        =   93
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox chWalkMode 
         Caption         =   "Modo caminata"
         Height          =   375
         Left            =   240
         TabIndex        =   122
         Top             =   7080
         Width           =   3015
      End
      Begin VB.Frame fRes 
         Caption         =   "Resolucion para"
         Height          =   1575
         Left            =   240
         TabIndex        =   114
         Top             =   5400
         Width           =   3375
         Begin VB.OptionButton optRes 
            Caption         =   "1152x864"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   124
            Top             =   1080
            Width           =   2535
         End
         Begin VB.OptionButton optRes 
            Caption         =   "1280x1024"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   116
            Top             =   720
            Width           =   2535
         End
         Begin VB.OptionButton optRes 
            Caption         =   "1024x768"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   115
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fVelocity 
         Caption         =   "Velocidad"
         Height          =   2535
         Left            =   240
         TabIndex        =   105
         Top             =   2640
         Width           =   3495
         Begin VB.TextBox txVelocity 
            Enabled         =   0   'False
            Height          =   285
            Left            =   360
            TabIndex        =   110
            Text            =   "1"
            Top             =   1890
            Width           =   1215
         End
         Begin VB.OptionButton optVel 
            Caption         =   "Elegir"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   109
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton optVel 
            Caption         =   "Ultra Rapido"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   108
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optVel 
            Caption         =   "Rapido"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   107
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton optVel 
            Caption         =   "Normal"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   106
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame fView 
         Caption         =   "Visiones"
         Height          =   2175
         Left            =   240
         TabIndex        =   94
         Top             =   240
         Width           =   3495
         Begin VB.CheckBox chVNPCs 
            Caption         =   "Ver NPCs"
            Height          =   375
            Left            =   120
            TabIndex        =   104
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1500
         End
         Begin VB.CheckBox chVLuces 
            Caption         =   "Ver Luces"
            Height          =   375
            Left            =   1920
            TabIndex        =   103
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1500
         End
         Begin VB.CheckBox chVCapa3 
            Caption         =   "Ver Capa 3"
            Height          =   375
            Left            =   120
            TabIndex        =   102
            Top             =   960
            Value           =   1  'Checked
            Width           =   1500
         End
         Begin VB.CheckBox chVCapa4 
            Caption         =   "Ver Capa 4"
            Height          =   375
            Left            =   120
            TabIndex        =   101
            Top             =   1320
            Width           =   1500
         End
         Begin VB.CheckBox chVBloq 
            Caption         =   "Ver Blockeos"
            Height          =   375
            Left            =   1920
            TabIndex        =   100
            Top             =   240
            Width           =   1500
         End
         Begin VB.CheckBox chVExit 
            Caption         =   "Ver Traslados"
            Height          =   375
            Left            =   1920
            TabIndex        =   99
            Top             =   600
            Width           =   1500
         End
         Begin VB.CheckBox chVTriggers 
            Caption         =   "Ver Triggers"
            Height          =   375
            Left            =   1920
            TabIndex        =   98
            Top             =   960
            Width           =   1500
         End
         Begin VB.CheckBox chVObjs 
            Caption         =   "Ver Objetos"
            Height          =   375
            Left            =   1920
            TabIndex        =   97
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1500
         End
         Begin VB.CheckBox chVCapa2 
            Caption         =   "Ver Capa 2"
            Height          =   375
            Left            =   120
            TabIndex        =   96
            Top             =   600
            Value           =   1  'Checked
            Width           =   1500
         End
         Begin VB.CheckBox chVCapa1 
            Caption         =   "Ver Capa 1"
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Value           =   1  'Checked
            Width           =   1500
         End
      End
   End
   Begin VB.Frame fMapa 
      Height          =   7695
      Left            =   120
      TabIndex        =   78
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdMusic 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   79
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNameMap 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   480
         Width           =   4095
      End
      Begin VB.CheckBox RESU 
         Caption         =   "Resu sin efecto"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox INVI 
         Caption         =   "Invi sin efecto"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox PK 
         Caption         =   "Insegura"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox Magia 
         Caption         =   "Magia sin efecto"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtMusicNum 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1920
         TabIndex        =   83
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMusic 
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   82
         Top             =   1200
         Width           =   855
      End
      Begin VB.ListBox terreno 
         Height          =   1230
         ItemData        =   "frmGrhView.frx":0020
         Left            =   120
         List            =   "frmGrhView.frx":0039
         TabIndex        =   81
         Top             =   3000
         Width           =   975
      End
      Begin VB.ListBox restringuir 
         Height          =   1230
         ItemData        =   "frmGrhView.frx":007B
         Left            =   1800
         List            =   "frmGrhView.frx":008E
         TabIndex        =   80
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label DimMap 
         Caption         =   "Dimensiones del Mapa : X = ; Y ="
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   4440
         Width           =   3735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Nombre del Mapa"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Número de Midi"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label16 
         Caption         =   "Terreno:"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Restringir para:"
         Height          =   255
         Left            =   1680
         TabIndex        =   89
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Frame fTriggers 
      Height          =   7695
      Left            =   120
      TabIndex        =   60
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdTrigSacarMap 
         Caption         =   "Sacar en Mapa"
         Height          =   315
         Left            =   240
         TabIndex        =   123
         Top             =   3960
         Width           =   2415
      End
      Begin VB.ListBox TrigList 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         ItemData        =   "frmGrhView.frx":00B7
         Left            =   240
         List            =   "frmGrhView.frx":00D6
         TabIndex        =   64
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chPutTrigger 
         Caption         =   "Insertar/Quitar"
         Height          =   240
         Left            =   240
         TabIndex        =   63
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdTrigInsertSelect 
         Caption         =   "Insertar en seleccion"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   62
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton cmdTrigSacarSelect 
         Caption         =   "Sacar en seleccion"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   61
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Para eliminar el Trigger, usar el trigger 0 o Click Derecho."
         Height          =   405
         Left            =   240
         TabIndex        =   65
         Top             =   2400
         Width           =   2415
      End
   End
   Begin VB.Frame fParticles 
      Height          =   7695
      Left            =   120
      TabIndex        =   57
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox chPutParticle 
         Caption         =   "Insertar/Quitar"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ListBox lstParticles 
         Height          =   2400
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame fAuto 
      Caption         =   "Automatizadores"
      Height          =   7695
      Left            =   120
      TabIndex        =   119
      Top             =   960
      Width           =   4335
      Begin VB.CommandButton cmdInsertCam 
         Caption         =   "Insertar Caminos"
         Height          =   375
         Left            =   120
         TabIndex        =   121
         Top             =   7200
         Width           =   4095
      End
      Begin VB.CheckBox chPutAuto 
         Caption         =   "Insertar recorrido"
         Height          =   375
         Left            =   120
         TabIndex        =   120
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fObj 
      Height          =   7695
      Left            =   120
      TabIndex        =   48
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdSacarObjSelect 
         Caption         =   "Sacar objs en seleccion"
         Enabled         =   0   'False
         Height          =   735
         Left            =   3000
         TabIndex        =   117
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdBloqArbol 
         Caption         =   "Bloquear Arboles"
         Height          =   615
         Left            =   3000
         TabIndex        =   112
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdMapArbol 
         Caption         =   "Mapear arboles"
         Height          =   615
         Left            =   3000
         TabIndex        =   111
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox picObj 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00000080&
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
         Height          =   3000
         Left            =   1800
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   75
         Top             =   4560
         Width           =   2400
      End
      Begin VB.CheckBox chPutObj 
         Caption         =   "Insertar/Quitar"
         Height          =   240
         Left            =   120
         TabIndex        =   53
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ListBox ObjList 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         ItemData        =   "frmGrhView.frx":0170
         Left            =   120
         List            =   "frmGrhView.frx":0177
         TabIndex        =   52
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox ObjCant 
         Height          =   285
         Left            =   960
         TabIndex        =   51
         Text            =   "1"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txSearchObj 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CheckBox ckCap3 
         Caption         =   "Añadir en capa 3"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Para eliminar el objeto utilice el Click Derecho."
         Height          =   885
         Left            =   1800
         TabIndex        =   56
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   3480
         Width           =   2655
      End
   End
   Begin VB.Frame fBlock 
      Height          =   7695
      Left            =   120
      TabIndex        =   39
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox chPutBlock 
         Caption         =   "Insertar/Quitar"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdBlockInsertSelect 
         Caption         =   "Insertar en seleccion"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   43
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdBlockSacSelect 
         Caption         =   "Sacar en seleccion"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   42
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cmdBlockInsertMap 
         Caption         =   "Insertar en Mapa"
         Height          =   315
         Left            =   240
         TabIndex        =   41
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdBlockInsertBord 
         Caption         =   "Insertar Bordes"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Para quitar bloqueos utilice el click derecho."
         Height          =   615
         Left            =   360
         TabIndex        =   45
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ObjClick As Integer


Private Sub chPutAuto_Click()
    If chPutAuto.value = vbChecked Then
        PutAuto = True
        ShowAuto = True
    Else
        PutAuto = False
        ShowAuto = False
    End If
End Sub

Private Sub chPutTraslado_Click()
    If chPutTraslado.value = Checked Then
        PutTrans = True
        ShowTrans = True
        chVExit.value = vbChecked
    Else
        PutTrans = False
    End If
End Sub

Private Sub chPutBlock_Click()
    If chPutBlock.value = Checked Then
        PutBlock = True
        chVBloq.value = vbChecked
        ShowBlocked = True
    Else
        PutBlock = False
    End If
End Sub

Private Sub chPutLight_Click()
    If chPutLight.value = Checked Then
        PutLight = True
        ShowLuces = True
        chVLuces.value = vbChecked
    Else
        PutLight = False
    End If
End Sub

Private Sub chPutNpc_Click()
    If chPutNpc.value = Checked Then
        PutNPC = True
        ShowNpcs = True
        chVNPCs.value = vbChecked
    Else
        PutNPC = False
    End If
End Sub

Private Sub chPutParticle_Click()
    If chPutParticle.value = Checked Then
        PutParticles = True
    Else
        PutParticles = False
    End If
End Sub

Private Sub chPutTrigger_Click()
    If chPutTrigger.value = Checked Then
        PutTrigger = True
        ShowTriggers = True
        chVTriggers.value = vbChecked
    Else
        PutTrigger = False
    End If
End Sub

Private Sub chVBloq_Click()
    If chVBloq.value = vbChecked Then
        ShowBlocked = True
    Else
        ShowBlocked = False
    End If
End Sub

Private Sub chVCapa1_Click()
    If chVCapa1.value = vbChecked Then
        ShowLayer1 = True
    Else
        ShowLayer1 = False
    End If
End Sub

Private Sub chVCapa2_Click()
    If chVCapa2.value = vbChecked Then
        ShowLayer2 = True
    Else
        ShowLayer2 = False
    End If
End Sub

Private Sub chVCapa3_Click()
    If chVCapa3.value = vbChecked Then
        ShowLayer3 = True
    Else
        ShowLayer3 = False
    End If
End Sub

Private Sub chVCapa4_Click()
    If chVCapa4.value = vbChecked Then
        bTecho = 1
    Else
        bTecho = 0
    End If
End Sub

Private Sub chVExit_Click()
    If chVExit.value = vbChecked Then
        ShowTrans = True
    Else
        ShowTrans = False
    End If
End Sub

Private Sub chVLuces_Click()
    If chVLuces.value = vbChecked Then
        ShowLuces = True
    Else
        ShowLuces = False
    End If
End Sub

Private Sub chVNPCs_Click()
    If chVNPCs.value = vbChecked Then
        ShowNpcs = True
    Else
        ShowNpcs = False
    End If
End Sub

Private Sub chVObjs_Click()
    If chVObjs.value = vbChecked Then
        ShowObjs = True
    Else
        ShowObjs = False
    End If
End Sub

Private Sub chVTriggers_Click()
    If chVTriggers.value = vbChecked Then
        ShowTriggers = True
    Else
        ShowTriggers = False
    End If
End Sub

Private Sub chWalkMode_Click()
    If chWalkMode.value = vbChecked Then
        ShowChar = True
        Engine.Char_Make 9990, 106, 10, 3, 1, 1, 2, 2, 2
    Else
        ShowChar = False
    End If
End Sub

Private Sub cmdBlockInsertBord_Click()
    Dim y As Integer
    Dim x As Integer

    modDeshacer.Deshacer_Add ""
        
    For y = 1 To MapInfo.dY
        For x = 1 To MapInfo.dX
            If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
                MapData(x, y).Blocked = 1
            End If
        Next x
    Next y
End Sub

Private Sub cmdBlockInsertMap_Click()
    Dim x1 As Integer, y1 As Integer
    
        modDeshacer.Deshacer_Add ""
        
    For x1 = 1 To 100
        For y1 = 1 To 100
            MapData(x1, y1).Blocked = 1
        Next y1
    Next x1
End Sub

Private Sub cmdBlockInsertSelect_Click()
    Dim x1 As Integer, y1 As Integer
    
    modDeshacer.Deshacer_Add ""
    
    For x1 = SelecOX To SelecDX
        For y1 = SelecOY To SelecDY
            MapData(x1, y1).Blocked = 1
        Next y1
    Next x1
End Sub

Private Sub cmdBlockSacSelect_Click()
    Dim x1 As Integer, y1 As Integer
    
    modDeshacer.Deshacer_Add ""
    For x1 = SelecOX To SelecDX
        For y1 = SelecOY To SelecDY
            MapData(x1, y1).Blocked = 0
        Next y1
    Next x1
End Sub

Private Sub cmdBloqArbol_Click()
    Dim x As Long
    Dim y As Long
    
    modDeshacer.Deshacer_Add ""
        
    For x = 1 To 100
        For y = 1 To 100
            CheckearCapa3 x, y
        
            If EsArbol(MapData(x, y).OBJInfo.ObjIndex) Then
                MapData(x, y).Blocked = 1
            End If
        Next y
    Next x
End Sub
Function EsArbol(ByVal Obj As Integer) As Boolean
    EsArbol = (Obj > 127 And Obj < 145)
End Function
Sub CheckearCapa3(ByVal x As Long, ByVal y As Long)
    Dim Obj As Integer
    If MapData(x, y).Graphic(3).GrhIndex <> 0 Then
        Select Case MapData(x, y).Graphic(3).GrhIndex
            Case 7000: Obj = 128
            Case 7001: Obj = 129
            Case 7002: Obj = 130
            Case 641: Obj = 131
            Case 642: Obj = 132
            Case 641: Obj = 133
            Case 644: Obj = 134
            Case 647: Obj = 135
            Case 7001: Obj = 136
            Case 7002: Obj = 137
            Case 7000: Obj = 138
            Case 7222: Obj = 139
            Case 7223: Obj = 140
            Case 7224: Obj = 141
            Case 7225: Obj = 142
            Case 7226: Obj = 143
            Case 735: Obj = 144
        End Select
        
        If Obj <> 0 Then
            MapData(x, y).OBJInfo.ObjIndex = Obj
            MapData(x, y).OBJInfo.amount = 1
            
            MapData(x, y).ObjGrh.GrhIndex = ObjData(Obj).GrhIndex
            Grh_Init MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex
        End If
    End If
End Sub

Private Sub cmdInsertNpcMap_Click()
    Dim cant As Long
    cant = Val(InputBox("Seleccione la cantidad de npcs que quiere que tenga el mapa"))
    
    If cant < 1 Then Exit Sub
    
    Dim i As Long, x As Long, y As Long
    Dim NPCIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
            
    NPCIndex = CInt(Val(ReadField(1, frmMode.lstNPC.Text, Asc("^"))))
    If NPCIndex = 0 Then Exit Sub
    
    Body = NpcData(NPCIndex).Body
    Head = NpcData(NPCIndex).Head
    Heading = NpcData(NPCIndex).Heading
            
    x = General_Random_Number(12, 88)
    y = General_Random_Number(12, 88)
        
    Deshacer_Add ""
        
    Do While cant > 0
            
        If MapData(x, y).Blocked = 0 And MapData(x, y).Trigger = 0 Then
            MapData(x, y).NPCIndex = NPCIndex
            Call Engine.Char_Make(NextOpenChar(), Body, Head, Heading, x, y, 0, 0, 0)
            
            cant = cant - 1
        End If
        
        x = General_Random_Number(12, 88)
        y = General_Random_Number(12, 88)
    Loop
        
End Sub

Private Sub cmdMapArbol_Click()
    Dim cant As Long
    cant = Val(InputBox("Seleccione la cantidad de arboles que quiere que tenga el mapa"))
    
    If cant < 1 Then Exit Sub
    
    Dim i As Long, x As Long, y As Long
    
    x = General_Random_Number(12, 88)
    y = General_Random_Number(12, 88)
        
    modDeshacer.Deshacer_Add ""
        
    Do While cant > 0
            
        If ((MapData(x, y).Graphic(1).GrhIndex > 6000 And _
            MapData(x, y).Graphic(1).GrhIndex < 6065) Or _
            (MapData(x, y).Graphic(1).GrhIndex > 13102 And _
            MapData(x, y).Graphic(1).GrhIndex < 13119)) And _
            MapData(x, y).OBJInfo.ObjIndex = 0 Then
            
            MapData(x, y).OBJInfo.amount = 1
            MapData(x, y).OBJInfo.ObjIndex = ObjClick
            
            MapData(x, y).Blocked = 1
            
            InitGrh MapData(x, y).ObjGrh, ObjData(ObjClick).GrhIndex
            
            cant = cant - 1
        End If
        
        x = General_Random_Number(12, 88)
        y = General_Random_Number(12, 88)
    Loop
        
End Sub

Private Sub cmdMusic_Click(Index As Integer)
    Select Case Index
        Case 0
            Audio.StopMidi
            Audio.PlayMIDI DirMidi & txtMusicNum.Text & ".mid"
        Case 1
            Audio.StopMidi
    End Select
End Sub

Private Sub cmdSacarObjSelect_Click()
    Dim x1 As Integer, y1 As Integer
    
    modDeshacer.Deshacer_Add ""
    For x1 = SelecOX To SelecDX
        For y1 = SelecOY To SelecDY
            MapData(x1, y1).ObjGrh.GrhIndex = 0
            MapData(x1, y1).OBJInfo.amount = 0
            MapData(x1, y1).OBJInfo.ObjIndex = 0
        Next y1
    Next x1
End Sub

Private Sub cmdShowUnion_Click()
    frmUnion.Show , Me
End Sub

Private Sub cmdSupInsertSelec_Click()
    Dim x1 As Integer, y1 As Integer
    modDeshacer.Deshacer_Add ""
        
    For x1 = SelecOX To SelecDX
        For y1 = SelecOY To SelecDY
            Dim SurfaceIndex As Integer
            If frmMode.lstSuperfices.Text = "" Then Exit Sub
            SurfaceIndex = CLng(ReadField(2, frmMode.lstSuperfices.Text, Asc("#")))
        
            If SupData(SurfaceIndex).Width > 0 Then
               
            Dim aux As Integer
            Dim dY As Integer
            Dim dX As Integer
                       
                dY = 0
                dX = 0
                
                aux = Val(SupData(SurfaceIndex).Grh) + _
                    (((y1 + dY) Mod SupData(SurfaceIndex).Height) * SupData(SurfaceIndex).Width) + ((x1 + dX) Mod SupData(SurfaceIndex).Width)
                  
                If MapData(x1, y1).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex <> aux Or MapData(x1, y1).Blocked <> frmMode.fSups.Visible Then
                    MapData(x1, y1).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex = aux
                    InitGrh MapData(x1, y1).Graphic(Val(frmMode.lstCapa.Text)), aux
                End If
            End If
        Next y1
    Next x1
End Sub

Private Sub cmdMenu_Click(Index As Integer)
    Limpiar
    
    Select Case Index
        Case 0
            fGrhView.Visible = True
            
        Case 1
            fNPCs.Visible = True
        
        Case 2
            fLuces.Visible = True
        
        Case 3
            fParticles.Visible = True
        
        Case 4
            fObj.Visible = True
        
        Case 5
            fBlock.Visible = True
        
        Case 6
            fTraslado.Visible = True
        
        Case 7
            fTriggers.Visible = True
            
        Case 8
            fOpc.Visible = True
            
        Case 9
            fMapa.Visible = True
            
        Case 10
            fAuto.Visible = True
            
    End Select
End Sub

Private Sub cmdSupSacSelec_Click()
    Dim x1 As Integer, y1 As Integer
    modDeshacer.Deshacer_Add ""
    For x1 = SelecOX To SelecDX
        For y1 = SelecOY To SelecDY
            MapData(x1, y1).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex = 0
        Next y1
    Next x1
End Sub


Private Sub chPutSup_Click()
    If chPutSup.value = Checked Then
        PutSurface = True
    Else
        PutSurface = False
    End If
End Sub

Private Sub cmdSupInsertMap_Click()
    If MsgBox("Esta seguro de insertar la superficie en todo el mapa?", vbYesNo) = vbNo Then Exit Sub
        
    modDeshacer.Deshacer_Add ""
    
    Dim y1 As Integer
    Dim x1 As Integer
    For x1 = 1 To MapInfo.dX
        For y1 = 1 To MapInfo.dY
            Dim SurfaceIndex As Integer
            SurfaceIndex = CLng(ReadField(2, frmMode.lstSuperfices.Text, Asc("#")))
            
            If SupData(SurfaceIndex).Width > 0 Then
               
                Dim aux As Integer
                Dim dY As Integer
                Dim dX As Integer
                             
                dY = 0
                dX = 0
                      
                frmMain.Change = True
                aux = Val(SupData(SurfaceIndex).Grh) + _
                    (((y1 + dY) Mod SupData(SurfaceIndex).Height) * SupData(SurfaceIndex).Width) + ((x1 + dX) Mod SupData(SurfaceIndex).Width)
                      
                If MapData(x1, y1).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex <> aux Or MapData(x1, y1).Blocked <> frmMode.fSups.Visible Then
                    MapData(x1, y1).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex = aux
                    InitGrh MapData(x1, y1).Graphic(Val(frmMode.lstCapa.Text)), aux
                End If
            End If
        Next y1
    Next x1
    frmMain.Change = True
End Sub

Public Sub Limpiar()
    PutBlock = False
    PutTrigger = False
    PutParticles = False
    PutObjs = False
    PutLight = False
    PutSurface = False
    PutNPC = False
    PutTrans = False
    PutAuto = False
    
    chPutLight.value = vbUnchecked
    chPutNpc.value = vbUnchecked
    chPutSup.value = vbUnchecked
    chPutBlock.value = vbUnchecked
    chPutObj.value = vbUnchecked
    chPutParticle.value = vbUnchecked
    chPutTrigger.value = vbUnchecked
    chPutTraslado.value = vbUnchecked
    chPutAuto.value = vbUnchecked
    
    fLuces.Visible = False
    fGrhView.Visible = False
    fNPCs.Visible = False
    fBlock.Visible = False
    fObj.Visible = False
    fParticles.Visible = False
    fTriggers.Visible = False
    fTraslado.Visible = False
    fMapa.Visible = False
    fOpc.Visible = False
    fAuto.Visible = False
End Sub


Private Sub cmdTrigInsertSelect_Click()
    modDeshacer.Deshacer_Add ""
    Dim x1 As Integer, y1 As Integer
    For x1 = SelecOX To SelecDX
        For y1 = SelecOY To SelecDY
            MapData(x1, y1).Trigger = frmMode.TrigList.ListIndex
        Next y1
    Next x1
End Sub

Private Sub cmdTrigSacarMap_Click()
    modDeshacer.Deshacer_Add ""
    Dim x1 As Integer, y1 As Integer
    For x1 = 1 To 100
        For y1 = 1 To 100
            MapData(x1, y1).Trigger = 0
        Next y1
    Next x1
End Sub

Private Sub cmdTrigSacarSelect_Click()
    modDeshacer.Deshacer_Add ""
    Dim x1 As Integer, y1 As Integer
    For x1 = SelecOX To SelecDX
        For y1 = SelecOY To SelecDY
            MapData(x1, y1).Trigger = 0
        Next y1
    Next x1
End Sub

Private Sub Form_Load()
    Limpiar
End Sub

Private Sub chPutObj_Click()
    If chPutObj.value = Checked Then
        PutObjs = True
        chVObjs.value = vbChecked
        ShowObjs = True
    Else
        PutObjs = False
    End If
End Sub

Private Sub lstNPC_Click()
    Dim GrhIndex As Long
    Dim rgb_temp(3) As Long
    Dim Body As Integer, Head As Integer
    Dim NPCIndex As Integer
    Dim DestRect As RECT
    
    rgb_temp(0) = -1
    rgb_temp(1) = -1
    rgb_temp(2) = -1
    rgb_temp(3) = -1
    
    DestRect.bottom = 200
    DestRect.Right = 160
    DestRect.Left = 0
    DestRect.Top = 0
    
    NPCIndex = CInt(Val(ReadField(1, frmMode.lstNPC.Text, Asc("^"))))
    Body = NpcData(NPCIndex).Body
    Head = NpcData(NPCIndex).Head
    
    If Body <> 0 Then Body = GrhData(BodyData(Body).Walk(3).GrhIndex).Frames(1)
    If Head <> 0 Then Head = HeadData(Head).Head(3).GrhIndex
    
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        If Head <> 0 Then Engine.Draw_GrhIndex Head, 16, 38 + BodyData(NpcData(NPCIndex).Body).HeadOffset.y, rgb_temp
        If Body <> 0 Then Engine.Draw_GrhIndex Body, 10, 10, rgb_temp
    D3DDevice.EndScene
    D3DDevice.Present DestRect, ByVal 0, picRenderNpc.hwnd, ByVal 0
End Sub

Private Sub lstSuperfices_Click()
On Error Resume Next

    Dim GrhIndex As Long
    Dim Ancho As Long
    Dim Alto As Long
    Dim i As Integer, j As Integer
    Dim DestRect As RECT
    Dim rgb_temp(3) As Long
    Dim Mosaico As Boolean
    Dim GrhIn As Long
    Dim Count As Integer
    GrhIn = CLng(ReadField(2, lstSuperfices.Text, Asc("#")))
    
    If SupData(GrhIn).Width > 0 Then
        Mosaico = True
    Else
        Mosaico = False
    End If
    
    Ancho = SupData(GrhIn).Width
    Alto = SupData(GrhIn).Height
    GrhIndex = SupData(GrhIn).Grh

    Label3.Caption = "Alto : " & Alto
    Label4.Caption = "Ancho : " & Ancho
    
    rgb_temp(0) = -1
    rgb_temp(1) = -1
    rgb_temp(2) = -1
    rgb_temp(3) = -1
    
    DestRect.bottom = 128
    DestRect.Right = 256
    DestRect.Left = 0
    DestRect.Top = 0
    
    If SupData(GrhIn).Capa > 0 Then
        lstCapa.ListIndex = SupData(GrhIn).Capa - 1
    End If
    
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    If Mosaico Then
        For i = 1 To CByte(Ancho)
            For j = 1 To CByte(Alto)
                GrhIndex = GrhData(GrhIndex).Frames(1)
                
                Engine.Draw_GrhIndex GrhIndex, (j - 1) * 32, (i - 1) * 32, rgb_temp()
                If Count < CInt(Val(Alto)) * CInt(Ancho) Then _
                     Count = Count + 1: GrhIndex = GrhIndex + 1
                     
               ' GrhIndex = GrhData(GrhIndex).Frames(1)
            Next
        Next
    Else
        Engine.Draw_GrhIndex GrhIndex, 0, 0, rgb_temp()
    End If
    GrhIndex = GrhIndex - Count
    D3DDevice.EndScene
    D3DDevice.Present DestRect, ByVal 0, frmMode.picSups.hwnd, ByVal 0
End Sub

Private Sub lstSuperfices_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Dim GrhIn As Long
        GrhIn = CLng(IIf(ReadField(2, lstSuperfices.Text, Asc("#")) = "", -1, ReadField(2, lstSuperfices.Text, Asc("#"))))
        
        If GrhIn = -1 Then Exit Sub
        SupData(GrhIn).name = ""
        lstSuperfices.List(lstSuperfices.ListIndex) = "-"
    End If
End Sub

Private Sub lstSuperfices_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MiMouse = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MiMouse = False
End Sub

Private Sub ObjList_Click()
    ObjClick = Val(ReadField(2, ObjList.List(ObjList.ListIndex), Asc("#")))
    
    Dim GrhIndex As Long
    Dim rgb_temp(3) As Long
    Dim DestRect As RECT
    
    rgb_temp(0) = -1
    rgb_temp(1) = -1
    rgb_temp(2) = -1
    rgb_temp(3) = -1
    
    DestRect.bottom = 200
    DestRect.Right = 160
    DestRect.Left = 0
    DestRect.Top = 0

    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        Engine.Draw_GrhIndex ObjData(ObjClick).GrhIndex, 8, 8, rgb_temp()
    D3DDevice.EndScene
    D3DDevice.Present DestRect, ByVal 0, picObj.hwnd, ByVal 0
End Sub

Private Sub INVI_Click()
    If INVI.value = vbChecked Then
        MapInfo.InviSinEfecto = 1
    Else
        MapInfo.InviSinEfecto = 0
    End If
End Sub

Private Sub Magia_Click()
    If Magia.value = vbChecked Then
        MapInfo.MagiaSinEfecto = 1
    Else
        MapInfo.MagiaSinEfecto = 0
    End If
End Sub

Public Sub optRes_Click(Index As Integer)
    frmMain.Top = 0
    frmMain.Left = 0
            
    Select Case Index
        Case 0
            frmMain.Width = 12030
            frmMain.Height = 9255
            
        Case 1
            frmMain.Width = 14520
            frmMain.Height = 14940
            
        Case 2
            frmMain.Width = 12690
            frmMain.Height = 12465
    End Select
    
    frmMode.Left = frmMain.Width
    frmMode.Top = 0
                
    frmMinimap.Left = frmMode.Width + frmMode.Left - frmMinimap.Width
    frmMinimap.Top = frmMode.Height
    
    UserPos.x = 50
    UserPos.y = 50
    
    Engine.Engine_Reset
End Sub

Public Sub optVel_Click(Index As Integer)
    Select Case Index
        Case 0
            ScrollPixelsPerFrameX = 32
            ScrollPixelsPerFrameY = 32
            txVelocity.Enabled = False
            
        Case 1
            ScrollPixelsPerFrameX = 40
            ScrollPixelsPerFrameY = 40
            txVelocity.Enabled = False
            
        Case 2
            ScrollPixelsPerFrameX = 106
            ScrollPixelsPerFrameY = 106
            txVelocity.Enabled = False
            
        Case 3
            txVelocity.Enabled = True
    End Select
End Sub

Private Sub PK_Click()
    If PK.value = vbChecked Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
End Sub

Private Sub RESU_Click()
    If RESU.value = vbChecked Then
        MapInfo.ResuSinEfecto = 1
    Else
        MapInfo.ResuSinEfecto = 0
    End If
End Sub

Private Sub terreno_Click()
    Select Case terreno.ListIndex
        Case 3  ' Ciudad
            MapInfo.terreno = eTerreno.Ciudad
        Case 4  ' Campo
            MapInfo.terreno = eTerreno.Campo
        Case 5  ' Dungeon
            MapInfo.terreno = eTerreno.Dungeon
        Case 0  ' Bosque
            MapInfo.terreno = eTerreno.Bosque
        Case 1  ' Nieve
            MapInfo.terreno = eTerreno.Nieve
        Case 2  ' Desierto
            MapInfo.terreno = eTerreno.Desierto
        Case Else
            MapInfo.terreno = eTerreno.Campo
    End Select
End Sub
Private Sub restringuir_Click()
    Select Case restringuir.ListIndex
        Case 0
            MapInfo.restringir = eRestringir.Nada
        Case 1
            MapInfo.restringir = eRestringir.Armada
        Case 2
            MapInfo.restringir = eRestringir.Caos
        Case 3
            MapInfo.restringir = eRestringir.Faccion
        Case 4
            MapInfo.restringir = eRestringir.Newbie
        Case Else
            MapInfo.restringir = eRestringir.Nada
    End Select
End Sub

Private Sub txtMusicNum_Change()
    MapInfo.Music = Val(txtMusicNum.Text)
End Sub

Private Sub txtNameMap_Change()
    MapInfo.name = txtNameMap.Text
End Sub

Private Sub txSearchObj_Change()
    Dim i As Long
    
    ObjList.Clear
    For i = 1 To UBound(ObjData())
        If InStr(UCase$(ObjData(i).name), UCase$(txSearchObj.Text)) > 0 Or txSearchObj.Text = "" Then
            ObjList.AddItem ObjData(i).name & " - #" & i
        End If
    Next
End Sub

Private Sub txSearchSurface_Change()
    Dim i As Long
    
    lstSuperfices.Clear
    For i = 0 To UBound(SupData())
        If InStr(UCase$(SupData(i).name), UCase$(txSearchSurface.Text)) > 0 Or txSearchSurface.Text = "" Then
            lstSuperfices.AddItem SupData(i).name & " - #" & i
        End If
    Next
End Sub

Private Sub txVelocity_Change()
    If Val(txVelocity.Text) < 1 Then
        txVelocity.Text = "1"
    ElseIf Val(txVelocity.Text) > 50 Then
        txVelocity.Text = "50"
    Else
        ScrollPixelsPerFrameX = Val(txVelocity.Text) * 32
        ScrollPixelsPerFrameY = Val(txVelocity.Text) * 32
    End If
End Sub
Function Insertar_Superficie(ByVal i As Integer, ByVal layer As Byte, ByVal destX As Integer, ByVal destY As Integer) As Boolean
    Dim tX As Integer, tY As Integer, despTile As Integer
        
    For tY = destY To destY + SupData(i).Height
        For tX = destX To destX + SupData(i).Width
            MapData(tX, tY).Graphic(layer).GrhIndex = CInt(Val(SupData(i).Grh) + despTile)
             
            despTile = despTile + 1
        Next
    Next
End Function

