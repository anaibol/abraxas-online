VERSION 5.00
Begin VB.Form frmSkills 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   1605
   ClientWidth     =   8940
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Guardar 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   6360
      TabIndex        =   42
      Top             =   4335
      Width           =   1095
   End
   Begin VB.Label Salir 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   41
      Top             =   4335
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos restantes:"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   5000
      TabIndex        =   40
      Top             =   3948
      Width           =   2055
   End
   Begin VB.Label Nombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Habilidades"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   3720
      TabIndex        =   39
      Top             =   255
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   1
      Left            =   600
      TabIndex        =   38
      Top             =   720
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   2
      Left            =   600
      TabIndex        =   37
      Top             =   1080
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   3
      Left            =   600
      TabIndex        =   36
      Top             =   1350
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   4
      Left            =   600
      TabIndex        =   35
      Top             =   1695
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   5
      Left            =   600
      TabIndex        =   34
      Top             =   2040
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   6
      Left            =   600
      TabIndex        =   33
      Top             =   2385
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   7
      Left            =   600
      TabIndex        =   32
      Top             =   2760
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   8
      Left            =   600
      TabIndex        =   31
      Top             =   3120
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   9
      Left            =   600
      TabIndex        =   30
      Top             =   3480
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   10
      Left            =   600
      TabIndex        =   29
      Top             =   3840
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   11
      Left            =   600
      TabIndex        =   28
      Top             =   4200
      Width           =   2205
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   12
      Left            =   4590
      TabIndex        =   27
      Top             =   750
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   1
      Left            =   3495
      TabIndex        =   26
      Top             =   750
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   2
      Left            =   3495
      TabIndex        =   25
      Top             =   1095
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   3
      Left            =   3495
      TabIndex        =   24
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   4
      Left            =   3495
      TabIndex        =   23
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   5
      Left            =   3495
      TabIndex        =   22
      Top             =   2145
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   6
      Left            =   3495
      TabIndex        =   21
      Top             =   2490
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   7
      Left            =   3495
      TabIndex        =   20
      Top             =   2850
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   8
      Left            =   3495
      TabIndex        =   19
      Top             =   3195
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   9
      Left            =   3495
      TabIndex        =   18
      Top             =   3540
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   10
      Left            =   3495
      TabIndex        =   17
      Top             =   3885
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Index           =   11
      Left            =   3495
      TabIndex        =   16
      Top             =   4200
      Width           =   555
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   12
      Left            =   7815
      TabIndex        =   15
      Top             =   750
      Width           =   555
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   3960
      Top             =   750
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   3960
      Top             =   1110
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   3180
      Top             =   1140
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   3960
      Top             =   1455
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   3180
      Top             =   1485
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   3960
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   3180
      Top             =   1830
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   3960
      Top             =   2145
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   3180
      Top             =   2175
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   3960
      Top             =   2490
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   3180
      Top             =   2520
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   12
      Left            =   3960
      Top             =   2835
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   13
      Left            =   3180
      Top             =   2865
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   14
      Left            =   3960
      Top             =   3180
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   15
      Left            =   3180
      Top             =   3210
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   16
      Left            =   3960
      Top             =   3540
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   17
      Left            =   3180
      Top             =   3540
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   18
      Left            =   3960
      Top             =   3885
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   19
      Left            =   3180
      Top             =   3885
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   20
      Left            =   3960
      Top             =   4230
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   21
      Left            =   3180
      Top             =   4230
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   22
      Left            =   8280
      Top             =   735
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   23
      Left            =   7380
      Top             =   825
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   24
      Left            =   8280
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   25
      Left            =   7380
      Top             =   1155
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   13
      Left            =   7815
      TabIndex        =   14
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   13
      Left            =   4590
      TabIndex        =   13
      Top             =   1095
      Width           =   2205
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   26
      Left            =   8280
      Top             =   1425
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   27
      Left            =   7380
      Top             =   1440
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   14
      Left            =   7815
      TabIndex        =   12
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   14
      Left            =   4590
      TabIndex        =   11
      Top             =   1440
      Width           =   2205
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   28
      Left            =   8280
      Top             =   1770
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   29
      Left            =   7380
      Top             =   1800
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   15
      Left            =   7815
      TabIndex        =   10
      Top             =   1785
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   15
      Left            =   4590
      TabIndex        =   9
      Top             =   1785
      Width           =   2205
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   30
      Left            =   8280
      Top             =   2115
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   31
      Left            =   7380
      Top             =   2160
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   16
      Left            =   7815
      TabIndex        =   8
      Top             =   2145
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   16
      Left            =   4590
      TabIndex        =   7
      Top             =   2130
      Width           =   2205
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   32
      Left            =   8280
      Top             =   2520
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   33
      Left            =   7380
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   17
      Left            =   7815
      TabIndex        =   6
      Top             =   2490
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   17
      Left            =   4590
      TabIndex        =   5
      Top             =   2475
      Width           =   2205
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   34
      Left            =   8280
      Top             =   2925
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   35
      Left            =   7380
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   18
      Left            =   7815
      TabIndex        =   4
      Top             =   2850
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   18
      Left            =   4590
      TabIndex        =   3
      Top             =   2820
      Width           =   2205
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3180
      Top             =   795
      Width           =   345
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Index           =   19
      Left            =   4590
      TabIndex        =   2
      Top             =   3165
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   19
      Left            =   7815
      TabIndex        =   1
      Top             =   3195
      Width           =   555
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   36
      Left            =   8280
      Top             =   3240
      Width           =   345
   End
   Begin VB.Image command1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   37
      Left            =   7380
      Top             =   3240
      Width           =   345
   End
   Begin VB.Label puntos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   6960
      TabIndex        =   0
      Top             =   3900
      Width           =   495
   End
End
Attribute VB_Name = "frmSkills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)

    Call Audio.Play(SND_CLICK)
    
    Dim indice
    If (Index And &H1) = 0 Then
        If Alocados > 0 Then
            indice = Index * 0.5 + 1
            If indice > NUMSKILLS Then indice = NUMSKILLS
            If Val(Text1(indice).Caption) < MAXSKILLPOINTS Then
                Text1(indice).Caption = Val(Text1(indice).Caption) + 1
                flags(indice) = flags(indice) + 1
                Alocados = Alocados - 1
            End If
        End If
    Else
        If Alocados < SkillPoints Then
            indice = Index * 0.5 + 1
            If Val(Text1(indice).Caption) > 0 And flags(indice) > 0 Then
                Text1(indice).Caption = Val(Text1(indice).Caption) - 1
                flags(indice) = flags(indice) - 1
                Alocados = Alocados + 1
            End If
        End If
    End If
    
    puntos.Caption = Alocados
End Sub

Private Sub Form_Load()

    Dim x As Long
    Dim y As Long
    Dim n As Long
    
    x = Width / Screen.TwipsPerPixelX
    y = Height / Screen.TwipsPerPixelY
    
    'set the corner angle by changing the value of 'n'
    n = 90
    
    SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, x, y, n, n), True
    
    'Nombres de los skills
    
    Dim L
    Dim i As Integer
    i = 1
    For Each L In label2
        L.Caption = SkillName(i)
        L.AutoSize = True
        i = i + 1
    Next
    i = 0
    
    'Flags para saber que skills se modificaron
    ReDim flags(1 To NUMSKILLS)
    
    'Cargamos el jpg correspondiente
    For i = 0 To NUMSKILLS + NUMSKILLS - 1
        If (i And &H1) = 0 Then
            Command1(i).Picture = LoadPicture(GrhPath & "BotónMás.jpg")
        Else
            Command1(i).Picture = LoadPicture(GrhPath & "BotónMenos.jpg")
        End If
    Next
    
    Call Make_Transparent_Form(hWnd, 230)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call Auto_Drag(hWnd)
    Else
        Unload Me
    End If
End Sub

Private Sub Guardar_Click()
On Error GoTo Error
    Dim skillChanges(NUMSKILLS) As Byte
    Dim i As Long

    For i = 1 To NUMSKILLS
        skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
        
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i
    
    Call WrItemodifySkills(skillChanges())
    
    If Alocados = 0 Then
        frmMain.imgAsignarSkill.Visible = False
    End If
    
    SkillPoints = Alocados
    Unload Me
Error:
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Salir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Salir.ForeColor = &H80FFFF
    Guardar.ForeColor = &H80FFFF
End Sub

Private Sub Guardar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Guardar.ForeColor = &H80FFFF
    Salir.ForeColor = &H80FFFF
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Guardar.ForeColor = &H80FFFF
    Salir.ForeColor = &H80FFFF
End Sub
