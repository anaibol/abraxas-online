VERSION 5.00
Begin VB.Form frmEstadisticas 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6990
   ClipControls    =   0   'False
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   1605
      TabIndex        =   27
      Top             =   810
      Width           =   240
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   1605
      TabIndex        =   26
      Top             =   1065
      Width           =   240
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   1605
      TabIndex        =   25
      Top             =   1335
      Width           =   240
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   1605
      TabIndex        =   24
      Top             =   1605
      Width           =   240
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   1605
      TabIndex        =   23
      Top             =   1890
      Width           =   240
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   5775
      TabIndex        =   22
      Top             =   825
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   5775
      TabIndex        =   21
      Top             =   1065
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   5775
      TabIndex        =   20
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   5775
      TabIndex        =   19
      Top             =   1590
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   5775
      TabIndex        =   18
      Top             =   1830
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   5775
      TabIndex        =   17
      Top             =   2085
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   5775
      TabIndex        =   16
      Top             =   2340
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   5775
      TabIndex        =   15
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   9
      Left            =   5775
      TabIndex        =   14
      Top             =   2895
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   10
      Left            =   5775
      TabIndex        =   13
      Top             =   3405
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   11
      Left            =   5775
      TabIndex        =   12
      Top             =   3660
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   12
      Left            =   5775
      TabIndex        =   11
      Top             =   3930
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   13
      Left            =   5775
      TabIndex        =   10
      Top             =   4185
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   14
      Left            =   5775
      TabIndex        =   9
      Top             =   4455
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   15
      Left            =   5775
      TabIndex        =   8
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   16
      Left            =   5775
      TabIndex        =   7
      Top             =   4935
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   17
      Left            =   5775
      TabIndex        =   6
      Top             =   5205
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   18
      Left            =   5775
      TabIndex        =   5
      Top             =   5475
      Width           =   360
   End
   Begin VB.Label Skills 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   19
      Left            =   5775
      TabIndex        =   4
      Top             =   5745
      Width           =   360
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1980
      TabIndex        =   3
      Top             =   2700
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   3330
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1980
      TabIndex        =   1
      Top             =   3030
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   930
      TabIndex        =   0
      Top             =   3630
      Width           =   825
   End
   Begin VB.Shape shpSkillsBar 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   1
      Left            =   5400
      Top             =   885
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   2
      Left            =   5400
      Top             =   1125
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   3
      Left            =   5400
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   4
      Left            =   5400
      Top             =   1635
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   5
      Left            =   5400
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   6
      Left            =   5400
      Top             =   2145
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   7
      Left            =   5400
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   8
      Left            =   5400
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   9
      Left            =   5400
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   10
      Left            =   5400
      Top             =   3465
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   11
      Left            =   5400
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   12
      Left            =   5400
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   13
      Left            =   5400
      Top             =   4245
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   14
      Left            =   5400
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   15
      Left            =   5400
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   16
      Left            =   5400
      Top             =   4995
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   17
      Left            =   5400
      Top             =   5265
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   18
      Left            =   5400
      Top             =   5535
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   19
      Left            =   5400
      Top             =   5805
      Width           =   1095
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonCerrar As clsGraphicalButton
Public LastPressed As clsGraphicalButton

Private Const ANCHO_BARRA As Byte = 73 'pixeles
Private Const BAR_LEFT_POS As Integer = 361 'pixeles

    Public Sub Iniciar_Labels()
    'Iniciamos los labels con los valores de los atributos y los skills
    Dim i As Integer
    Dim Ancho As Integer
    
    For i = 1 To NUMATRIBUTOS
        Atri(i).Caption = UserAtributos(i)
    Next
    
    For i = 1 To NUMSKILLS
        Skills(i).Caption = UserSkills(i)
        Ancho = IIf(PorcentajeSkills(i) = 0, ANCHO_BARRA, (100 - PorcentajeSkills(i)) * 0.01 * ANCHO_BARRA)
        shpSkillsBar(i).Width = Ancho
        shpSkillsBar(i).Left = BAR_LEFT_POS + ANCHO_BARRA - Ancho
    Next
    
    With UserEstadisticas
        Label6(0).Caption = .Matados
        Label6(1).Caption = .Muertes
        Label6(2).Caption = .NpcsMatados
        Label6(3).Caption = .Clase
    End With
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

    Call Make_Transparent_Form(hWnd, 225)
       
    Picture = LoadPicture(GrhPath & "Estadísticas.jpg")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call Auto_Drag(hWnd)
    Else
        Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
