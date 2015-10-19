VERSION 5.00
Begin VB.Form frmCustomKeys 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Guardar y Salir"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   45
      Top             =   6960
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Restaurar Predeterminados"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   44
      Top             =   6360
      Width           =   3375
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3780
      Left            =   2880
      TabIndex        =   4
      Top             =   2400
      Width           =   3735
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   20
         Left            =   2280
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   360
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   12
         Left            =   2280
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   2880
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   13
         Left            =   2280
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   3240
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   26
         Left            =   2280
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   2520
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   21
         Left            =   2280
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   25
         Left            =   2280
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   2160
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   24
         Left            =   2280
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   1800
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   23
         Left            =   2280
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   1440
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   22
         Left            =   2280
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Modo Seguro"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   53
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Seg. de Resucitación"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   50
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Capturar Pantalla"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   46
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Macro Trabajo"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Macro Hechizos"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Meditar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar Opciones"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar/Ocultar FPS"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Hablar"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   6240
      Width           =   2775
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   18
         Left            =   1440
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   360
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   19
         Left            =   1440
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar al Clan"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar a Todos"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3375
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   14
         Left            =   1080
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   1800
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   9
         Left            =   1080
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   17
         Left            =   1080
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2880
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   16
         Left            =   1080
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2520
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   15
         Left            =   1080
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2160
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1440
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   10
         Left            =   1080
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1080
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Atacar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Usar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Tirar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Ocultar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Robar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Domar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Equipar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Agarrar"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Opciones Personales"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1575
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   3735
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   360
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1080
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar/Ocultar Nombres"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Corregir Posicion"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Activar/Desactivar Musica"
         BeginProperty Font 
            Name            =   "FrizQuadrata BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Movimiento"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2415
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1440
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1080
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Derecha"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Izquierda"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Abajo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Arriba"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Label LabelControls 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración de controles"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   59
      Top             =   195
      Width           =   3495
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call CustomKeys.LoadDefaults
    
    Dim i As Long
    
    For i = 1 To CustomKeys.Count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub

Private Sub Command2_Click()
    Dim i As Long
    
    For i = 1 To CustomKeys.Count
        If LenB(Text1(i).Text) = 0 Then
            Call MsgBox("Hay una o mas teclas no validas, por favor verifique.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Abraxas")
            Exit Sub
        End If
    Next i
    
    Unload Me
End Sub

Private Sub Form_Load()

    Dim x As Long
    Dim y As Long
    Dim n As Long
    
    x = Width / Screen.TwipsPerPixelX
    y = Height / Screen.TwipsPerPixelY
    
    'set the corner angle by changing the value of 'n'
    n = 25
    
    SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, x, y, n, n), True
    
    Dim i As Long
        
    For i = 1 To CustomKeys.Count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub

Private Sub text1_Click(Index As Integer)
    Text1(Index).Enabled = True
    Text1(Index).ForeColor = vbRed
End Sub

Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Text1(Index).Enabled = True
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If Text1(Index).Enabled = False Then
        Exit Sub
    End If
    
    Dim i As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then
        Exit Sub
    End If
    
    'If key is not valid, we exit
    
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    For i = 1 To CustomKeys.Count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).Text = vbNullString 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Text1(Index).Enabled = False Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not Text1(Index).Enabled Then
        Exit Sub
    End If
    
    Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub
