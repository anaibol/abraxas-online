VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Abraxas"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   300
   ScaleMode       =   2  'Point
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton RandomName 
      Appearance      =   0  'Flat
      Caption         =   "?"
      Height          =   405
      Left            =   5880
      TabIndex        =   24
      Top             =   720
      Width           =   435
   End
   Begin VB.PictureBox picPJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   7125
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   6435
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   23
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   8055
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   22
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   6840
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   19
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Timer tAnimacion 
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtMail 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   4230
      TabIndex        =   5
      Top             =   2860
      Width           =   2400
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      MaxLength       =   15
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.ComboBox lstClase 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   345
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   480
      List            =   "frmCrearPersonaje.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   1860
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   345
      ItemData        =   "frmCrearPersonaje.frx":0004
      Left            =   2520
      List            =   "frmCrearPersonaje.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   1785
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   345
      ItemData        =   "frmCrearPersonaje.frx":0021
      Left            =   4440
      List            =   "frmCrearPersonaje.frx":0023
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   1620
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   7650
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   21
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   7245
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   20
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4230
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3750
      Width           =   2400
   End
   Begin VB.TextBox txtConfirmPasswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4230
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   4575
      Width           =   2400
   End
   Begin VB.Shape txtConfirmPasswdBorder 
      BorderColor     =   &H00C0FFFF&
      Height          =   330
      Left            =   4215
      Top             =   4560
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Shape txtPasswdBorder 
      BorderColor     =   &H00C0FFFF&
      Height          =   340
      Left            =   4215
      Top             =   3735
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Shape txtMailBorder 
      BorderColor     =   &H00C0FFFF&
      Height          =   420
      Left            =   4215
      Top             =   2850
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Shape NameTxtBorder 
      BorderColor     =   &H00C0FFFF&
      Height          =   400
      Left            =   3105
      Top             =   705
      Width           =   2680
   End
   Begin VB.Shape HeadSelectionBorder 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   405
      Left            =   7245
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   7125
      Picture         =   "frmCrearPersonaje.frx":0025
      Top             =   3735
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   7500
      Picture         =   "frmCrearPersonaje.frx":0337
      Top             =   3735
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   0
      Left            =   6360
      Picture         =   "frmCrearPersonaje.frx":0649
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   1
      Left            =   8280
      Picture         =   "frmCrearPersonaje.frx":095B
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label AgregadoAgilidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   2295
      TabIndex        =   16
      Top             =   3340
      Width           =   450
   End
   Begin VB.Label AgregadoFuerza 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   2295
      TabIndex        =   15
      Top             =   2960
      Width           =   450
   End
   Begin VB.Label AgregadoInteligencia 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   2295
      TabIndex        =   14
      Top             =   3740
      Width           =   450
   End
   Begin VB.Label AgregadoCarisma 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   2295
      TabIndex        =   13
      Top             =   4155
      Width           =   450
   End
   Begin VB.Label AgregadoConstitucion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   2295
      TabIndex        =   12
      Top             =   4560
      Width           =   450
   End
   Begin VB.Image lblSumaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   3570
      MouseIcon       =   "frmCrearPersonaje.frx":0C6D
      MousePointer    =   99  'Custom
      Top             =   2970
      Width           =   300
   End
   Begin VB.Image lblSumaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   4
      Left            =   3570
      MouseIcon       =   "frmCrearPersonaje.frx":0DBF
      MousePointer    =   99  'Custom
      Top             =   4665
      Width           =   300
   End
   Begin VB.Image lblSumaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   3
      Left            =   3560
      MouseIcon       =   "frmCrearPersonaje.frx":0F11
      MousePointer    =   99  'Custom
      Top             =   4260
      Width           =   300
   End
   Begin VB.Image lblSumaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   2
      Left            =   3570
      MouseIcon       =   "frmCrearPersonaje.frx":1063
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   300
   End
   Begin VB.Image lblSumaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   3570
      MouseIcon       =   "frmCrearPersonaje.frx":11B5
      MousePointer    =   99  'Custom
      Top             =   3400
      Width           =   300
   End
   Begin VB.Image lblRestaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   2700
      MouseIcon       =   "frmCrearPersonaje.frx":1307
      MousePointer    =   99  'Custom
      Top             =   2985
      Width           =   300
   End
   Begin VB.Image lblRestaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   4
      Left            =   2685
      MouseIcon       =   "frmCrearPersonaje.frx":1459
      MousePointer    =   99  'Custom
      Top             =   4665
      Width           =   300
   End
   Begin VB.Image lblRestaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   3
      Left            =   2685
      MouseIcon       =   "frmCrearPersonaje.frx":15AB
      MousePointer    =   99  'Custom
      Top             =   4260
      Width           =   300
   End
   Begin VB.Image lblRestaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   2
      Left            =   2685
      MouseIcon       =   "frmCrearPersonaje.frx":16FD
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   300
   End
   Begin VB.Image lblRestaAtributos 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   2700
      MouseIcon       =   "frmCrearPersonaje.frx":184F
      MousePointer    =   99  'Custom
      Top             =   3400
      Width           =   300
   End
   Begin VB.Label AtributosLibres 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   5200
      Width           =   375
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   1
      Left            =   6120
      MouseIcon       =   "frmCrearPersonaje.frx":19A1
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   1140
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   0
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":1AF3
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Height          =   315
      Index           =   3
      Left            =   3015
      TabIndex        =   11
      Top             =   4140
      Width           =   375
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Height          =   315
      Index           =   2
      Left            =   3015
      TabIndex        =   10
      Top             =   3750
      Width           =   375
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Height          =   315
      Index           =   4
      Left            =   3015
      TabIndex        =   9
      Top             =   4515
      Width           =   375
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Height          =   315
      Index           =   1
      Left            =   3015
      TabIndex        =   8
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Height          =   315
      Index           =   0
      Left            =   3015
      TabIndex        =   0
      Top             =   2970
      Width           =   375
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Conectando As Boolean

Private Dir As eHeading

Private currentGrh As Long

Private Const HUMANO_H_PRIMER_CABEZA As Integer = 1
Private Const HUMANO_H_ULTIMA_CABEZA As Integer = 40 'En verdad es hasta la 51, pero como son muchas estas las dejamos no sEleccionables
Private Const HUMANO_H_CUERPO_DESNUDO As Integer = 21

Private Const ELFO_H_PRIMER_CABEZA As Integer = 101
Private Const ELFO_H_ULTIMA_CABEZA As Integer = 122
Private Const ELFO_H_CUERPO_DESNUDO As Integer = 210

Private Const DROW_H_PRIMER_CABEZA As Integer = 201
Private Const DROW_H_ULTIMA_CABEZA As Integer = 221
Private Const DROW_H_CUERPO_DESNUDO As Integer = 32

Private Const ENANO_H_PRIMER_CABEZA As Integer = 301
Private Const ENANO_H_ULTIMA_CABEZA As Integer = 319
Private Const ENANO_H_CUERPO_DESNUDO As Integer = 53

Private Const GNOMO_H_PRIMER_CABEZA As Integer = 401
Private Const GNOMO_H_ULTIMA_CABEZA As Integer = 416
Private Const GNOMO_H_CUERPO_DESNUDO As Integer = 222
'
Private Const HUMANO_M_PRIMER_CABEZA As Integer = 70
Private Const HUMANO_M_ULTIMA_CABEZA As Integer = 89
Private Const HUMANO_M_CUERPO_DESNUDO As Integer = 39

Private Const ELFO_M_PRIMER_CABEZA As Integer = 170
Private Const ELFO_M_ULTIMA_CABEZA As Integer = 188
Private Const ELFO_M_CUERPO_DESNUDO As Integer = 259

Private Const DROW_M_PRIMER_CABEZA As Integer = 270
Private Const DROW_M_ULTIMA_CABEZA As Integer = 288
Private Const DROW_M_CUERPO_DESNUDO As Integer = 40

Private Const ENANO_M_PRIMER_CABEZA As Integer = 370
Private Const ENANO_M_ULTIMA_CABEZA As Integer = 384
Private Const ENANO_M_CUERPO_DESNUDO As Integer = 60

Private Const GNOMO_M_PRIMER_CABEZA As Integer = 470
Private Const GNOMO_M_ULTIMA_CABEZA As Integer = 484
Private Const GNOMO_M_CUERPO_DESNUDO As Integer = 260

Public SkillPoints As Byte

Private Sub boton_Click(Index As Integer)
    Call Audio.Play(SND_CLICK)
    
    Select Case Index
    
        Case 0
            If LenB(NameTxt.Text) < 1 Then
                MsgBox "No elegiste tu nuevo nombre."
                NameTxtBorder.BorderColor = &H80&
                NameTxt.SetFocus
                Exit Sub
            End If
            
            If Len(NameTxt.Text) < 3 Or Len(NameTxt.Text) > 15 Then
                MsgBox "Tu nombre debe debe contener 3 o más carácteres."
                NameTxtBorder.BorderColor = &H80&
                NameTxt.SetFocus
                Exit Sub
            End If
            
            'Validamos los datos del user
            Dim loopc As Byte
            Dim CharAscii As Integer
            
            For loopc = 1 To Len(NameTxt.Text)
                CharAscii = Asc(mid$(NameTxt.Text, loopc, 1))
                If Not LegalChar(CharAscii) Then
                    MsgBox "El nombre que elegiste es inválido porque contiene el carácter " & Chr(34) & Chr$(CharAscii) & Chr(34) & ", que no está permitido."
                    NameTxtBorder.BorderColor = &H80&
                    NameTxt.SetFocus
                    Exit Sub
                End If
            Next loopc
            
            If Right$(NameTxt.Text, 1) = " " Then
                UserName = RTrim$(NameTxt.Text)
                MsgBox "Fueron removidos los espacios al final del nombre."
            End If
            
            If lstRaza.ListIndex < 0 Then
                MsgBox "No elegiste tu raza."
                lstRaza.SetFocus
                Exit Sub
            End If
            
            If lstClase.ListIndex < 0 Then
                MsgBox "No elegiste tu clase."
                lstClase.SetFocus
                Exit Sub
            End If
            
            If lstGenero.ListIndex < 0 Then
                MsgBox "No elegiste tu género."
                lstGenero.SetFocus
                Exit Sub
            End If
            
            If AtributosLibres > 0 Then
                If AtributosLibres = 36 Then
                    MsgBox "No has asignado tus atributos."
                Else
                    MsgBox "No has terminado de asignar tus atributos."
                End If
                Exit Sub
            End If
            
            If LenB(txtMail.Text) < 1 Then
                MsgBox "No has introducido tu dirección de correo electrónico."
                txtMailBorder.BorderColor = &H80&
                txtMail.SetFocus
                Exit Sub
            End If
            
            If Not CheckMailString(txtMail.Text) Then
                MsgBox "La dirección de correo electrónico es inválida."
                txtMailBorder.BorderColor = &H80&
                txtMail.SetFocus
                Exit Sub
            End If
                            
            If Len(txtPasswd.Text) < 6 Then
                If LenB(txtPasswd.Text) < 1 Then
                    MsgBox "No elegiste tu clave."
                Else
                    MsgBox "El largo de tu clave debe ser de al menos 6 carácteres."
                End If
                
                txtPasswdBorder.BorderColor = &H80&
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            
            If Len(txtConfirmPasswd.Text) < 6 Then
                If LenB(txtConfirmPasswd.Text) < 1 Then
                    MsgBox "Repite tu clave para confirmarla."
                Else
                    MsgBox "Las claves no coinciden."
                End If
                
                txtConfirmPasswdBorder.BorderColor = &H80&
                txtConfirmPasswd.SetFocus
                Exit Sub
            End If
            
            If txtConfirmPasswd.Text <> txtPasswd.Text Then
                MsgBox "Las claves no coinciden. Debes repetir tu clave elegida para confirmarla."
                
                txtConfirmPasswdBorder.BorderColor = &H80&
                txtConfirmPasswd.SetFocus
                Exit Sub
            End If
    
            MousePointer = 11
    
            UserName = NameTxt.Text
            UserPassword = txtPasswd.Text
            UserEmail = txtMail.Text
            UserRaza = lstRaza.ListIndex + 1
            UserSexo = lstGenero.ListIndex + 1
            UserClase = lstClase.ListIndex + 1
            
            For loopc = 1 To NUMATRIBUTOS
                UserAtributos(loopc) = Val(lblAtributos(loopc - 1).Caption)
            Next loopc
            
            EstadoLogin = Creado
                  
            If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
                DoEvents
            End If
            
            frmMain.Socket1.Startup
                
            frmMain.Socket1.HostName = ServerIP
            frmMain.Socket1.RemotePort = ServerPort
            
            frmMain.Socket1.Connect
            
            Conectando = True
                            
        Case 1
            If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
                DoEvents
            ElseIf Conectando Then
                frmMain.Socket1.Abort
            End If
            
            frmConnect.MousePointer = vbDefault
            
            Unload Me
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
            
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
    End If
End Sub

Private Sub lstClase_Click()
    Call Audio.Play(SND_CLICK)
End Sub

Private Sub lstGenero_Click()
    Call Audio.Play(SND_CLICK)
    
    UserSexo = lstGenero.ListIndex + 1
    Call DarCuerpoYCabeza
End Sub

Private Sub lstRaza_Click()
    Call Audio.Play(SND_CLICK)

    UserRaza = lstRaza.ListIndex + 1
    Call DarCuerpoYCabeza
    
    Select Case (lstRaza.List(lstRaza.ListIndex))
        Case Is = "Humano"
            AgregadoFuerza.Caption = "+1"
            AgregadoAgilidad.Caption = "+1"
            AgregadoInteligencia.Caption = vbNullString
            AgregadoCarisma.Caption = vbNullString
            AgregadoConstitucion.Caption = "+2"
            
        Case Is = "Elfo"
            AgregadoFuerza.Caption = "-1"
            AgregadoAgilidad.Caption = "+3"
            AgregadoInteligencia.Caption = "+2"
            AgregadoCarisma.Caption = "+2"
            AgregadoConstitucion.Caption = "+1"
            
            AgregadoFuerza.ForeColor = &HC0&
        Case Is = "Elfo Oscuro"
            AgregadoFuerza.Caption = "+2"
            AgregadoAgilidad.Caption = "+3"
            AgregadoInteligencia.Caption = "+2"
            AgregadoCarisma.Caption = "-3"
            AgregadoConstitucion.Caption = vbNullString
            
            AgregadoCarisma.ForeColor = &HC0&
        Case Is = "Enano"
            AgregadoFuerza.Caption = "+3"
            AgregadoAgilidad.Caption = vbNullString
            AgregadoInteligencia.Caption = "-2"
            AgregadoCarisma.Caption = "-2"
            AgregadoConstitucion.Caption = "+3"
            
            AgregadoInteligencia.ForeColor = &HC0&
            AgregadoCarisma.ForeColor = &HC0&
        Case Is = "Gnomo"
            AgregadoFuerza.Caption = "-2"
            AgregadoAgilidad.Caption = "+3"
            AgregadoInteligencia.Caption = "+4"
            AgregadoCarisma.Caption = "+1"
            AgregadoConstitucion.Caption = vbNullString
            
            AgregadoFuerza.ForeColor = &HC0&
    End Select

End Sub

Private Sub NameTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call boton_Click(0)
    ElseIf KeyCode = vbKeyEscape Then
        Call boton_Click(1)
    End If
End Sub

Private Sub DirPJ_Click(Index As Integer)
    Select Case Index
        Case 0
            Dir = CheckDir(Dir + 1)
        Case 1
            Dir = CheckDir(Dir - 1)
    End Select
    
    Call UpdateHeadSelection
End Sub

Private Sub Form_Load()

    'If Not ChangeResolution And Not ResolucionActual Then
        'frmCrearPersonaje.BorderStyle = 3
        'frmCrearPersonaje.Icon = frmMain.Icon
        'frmCrearPersonaje.Caption = frmCrearPersonaje.Caption
    'End If
    
    Picture = LoadPicture(GrhPath & "CrearPersonaje.jpg")
    
    Dim i As Integer
    lstClase.Clear
    For i = LBound(ListaClases) To UBound(ListaClases)
        lstClase.AddItem ListaClases(i)
    Next i
    
    lstRaza.Clear
    
    For i = LBound(ListaRazas()) To UBound(ListaRazas())
        lstRaza.AddItem ListaRazas(i)
    Next i
    
    'lstClase.ListIndex = 1
    'lstRaza.ListIndex = 1
    'lstGenero.ListIndex = 1
    
    'UserRaza = lstRaza.ListIndex + 1
    'UserSexo = lstGenero.ListIndex + 1

    'Call DrawImageInPicture(picPJ, Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    Dir = SOUTH
End Sub

Private Sub HeadPJ_Click(Index As Integer)
    Select Case Index
        Case 0
            UserHead = CheckCabeza(UserHead + 1)
        Case 1
            UserHead = CheckCabeza(UserHead - 1)
    End Select
    
    Call UpdateHeadSelection
End Sub

Private Sub lblRestaAtributos_Click(Index As Integer)
    Call Audio.Play(SND_CLICK)

    If lblAtributos(Index).Caption > 8 Then
        lblAtributos(Index).Caption = Val(lblAtributos(Index).Caption) - 1
        AtributosLibres.Caption = Val(AtributosLibres.Caption) + 1
    End If
End Sub

Private Sub lblSumaAtributos_Click(Index As Integer)
    Call Audio.Play(SND_CLICK)
    
    If AtributosLibres.Caption > 0 And lblAtributos(Index).Caption < 18 Then
        lblAtributos(Index).Caption = Val(lblAtributos(Index).Caption) + 1
        AtributosLibres.Caption = Val(AtributosLibres.Caption) - 1
    End If
End Sub

Private Sub NameTxt_Change()
    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    Dim i As Long
    Dim tempstr As String
    Dim CharAscii As Integer
    
    For i = 1 To Len(NameTxt.Text)
        CharAscii = Asc(mid$(NameTxt.Text, i, 1))
        
        If (CharAscii > 64 And CharAscii < 91) Or (CharAscii > 96 And CharAscii < 123) Or CharAscii = 32 Then
            tempstr = tempstr & Chr$(CharAscii)
        End If
    Next i
    
    tempstr = StrConv(tempstr, vbProperCase)

    If tempstr <> NameTxt.Text Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        NameTxt.Text = tempstr
    End If

    If LenB(NameTxt.Text) > 0 Then
        'If LTrim$(UCase$(Left$(NameTxt.Text, 1)) & LCase$(Right$(NameTxt.Text, Len(NameTxt.Text) - 1))) <> NameTxt.Text Then
        'NameTxt.Text = LTrim$(UCase$(Left$(NameTxt.Text, 1)) & LCase$(Right$(NameTxt.Text, Len(NameTxt.Text) - 1)))
        'End If
        
        NameTxt.SelStart = Len(NameTxt.Text)
    End If
End Sub

Private Sub picHead_Click(Index As Integer)
    'No se mueve si clickea al medio
    If Index = 2 Then
        Exit Sub
    End If
    
    Dim Counter As Integer
    Dim Head As Integer
    
    Head = UserHead
    
    If Index > 2 Then
        For Counter = Index - 2 To 1 Step -1
            Head = CheckCabeza(Head + 1)
        Next Counter
    Else
        For Counter = 2 - Index To 1 Step -1
            Head = CheckCabeza(Head - 1)
        Next Counter
    End If
    
    UserHead = Head
    
    Call UpdateHeadSelection
End Sub

Private Sub DarCuerpoYCabeza()

    Dim bVisible As Boolean
    Dim PicIndex As Integer
    Dim LineIndex As Integer
    
    Select Case UserSexo
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_H_PRIMER_CABEZA
                    UserBody = HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_H_PRIMER_CABEZA
                    UserBody = ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_H_PRIMER_CABEZA
                    UserBody = DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_H_PRIMER_CABEZA
                    UserBody = ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_H_PRIMER_CABEZA
                    UserBody = GNOMO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_M_PRIMER_CABEZA
                    UserBody = HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_M_PRIMER_CABEZA
                    UserBody = ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_M_PRIMER_CABEZA
                    UserBody = DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_M_PRIMER_CABEZA
                    UserBody = ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_M_PRIMER_CABEZA
                    UserBody = GNOMO_M_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
        Case Else
            UserHead = 0
            UserBody = 0
    End Select
    
    bVisible = UserHead > 0 And UserBody > 0
    
    picPJ.Visible = bVisible
    HeadPJ(0).Visible = bVisible
    HeadPJ(1).Visible = bVisible
    DirPJ(0).Visible = bVisible
    DirPJ(1).Visible = bVisible
    
    For PicIndex = 0 To 4
        picHead(PicIndex).Visible = bVisible
    Next PicIndex
    
    HeadSelectionBorder.Visible = bVisible
    
    If bVisible Then
        Call UpdateHeadSelection
    End If
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    
    If currentGrh > 0 Then
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames) + 90
    End If
    
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

    Select Case UserSexo
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    If Head > HUMANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
                        CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Elfo
                    If Head > ELFO_H_ULTIMA_CABEZA Then
                        CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < ELFO_H_PRIMER_CABEZA Then
                        CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.ElfoOscuro
                    If Head > DROW_H_ULTIMA_CABEZA Then
                        CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < DROW_H_PRIMER_CABEZA Then
                        CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Enano
                    If Head > ENANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < ENANO_H_PRIMER_CABEZA Then
                        CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Gnomo
                    If Head > GNOMO_H_ULTIMA_CABEZA Then
                        CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
                        CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case Else
                    UserRaza = lstRaza.ListIndex + 1
                    CheckCabeza = CheckCabeza(Head)
            End Select
            
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    If Head > HUMANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
                        CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Elfo
                    If Head > ELFO_M_ULTIMA_CABEZA Then
                        CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < ELFO_M_PRIMER_CABEZA Then
                        CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.ElfoOscuro
                    If Head > DROW_M_ULTIMA_CABEZA Then
                        CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < DROW_M_PRIMER_CABEZA Then
                        CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Enano
                    If Head > ENANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < ENANO_M_PRIMER_CABEZA Then
                        CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Gnomo
                    If Head > GNOMO_M_ULTIMA_CABEZA Then
                        CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
                        CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case Else
                    UserRaza = lstRaza.ListIndex + 1
                    CheckCabeza = CheckCabeza(Head)
            End Select
        Case Else
            UserSexo = lstGenero.ListIndex + 1
            CheckCabeza = CheckCabeza(Head)
    End Select
End Function

Private Function CheckDir(ByRef Dir As eHeading) As eHeading

    If Dir > eHeading.WEST Then
        Dir = eHeading.NORTH
    End If
    
    If Dir < eHeading.NORTH Then
        Dir = eHeading.WEST
    End If
    
    CheckDir = Dir
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    
    If currentGrh > 0 Then
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames) + 90
    End If
    
End Function

Private Sub UpdateHeadSelection()
    Dim Head As Integer
    
    Head = UserHead
    Call DrawHead(Head, 2)
    
    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 3)
    
    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 4)
    
    Head = UserHead
    
    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 1)
    
    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 0)
End Sub

Private Sub RandomName_Click()
    If Not MainTimer.Check(TimersIndex.RandomName) Then
        Exit Sub
    End If
    
    EstadoLogin = BuscandoNombre
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If

    frmMain.Socket1.Startup
            
    frmMain.Socket1.HostName = ServerIP
    frmMain.Socket1.RemotePort = ServerPort

    frmMain.Socket1.Connect
End Sub

Private Sub tAnimacion_Timer()
    Dim SR As RECT
    Dim Grh As Long
    Dim x As Long
    Dim y As Long
    Static Frame As Byte
    
    If currentGrh = 0 Then
        Exit Sub
    End If
    
    UserHead = CheckCabeza(UserHead)
    
    Frame = Frame + 1
    
    If Frame >= GrhData(currentGrh).NumFrames Then
        Frame = 1
    End If
    
    Call DrawImageInPicture(picPJ, Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    
    Grh = GrhData(currentGrh).Frames(Frame)
    
    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .PixelWidth
        SR.Bottom = SR.Top + .PixelHeight
        
        x = picPJ.Width * 0.5 - .PixelWidth * 0.5 + 4
        y = picPJ.Height - .PixelHeight + 10
        
        Call DrawTransparentGrhtoHdc(picPJ.hdc, x, y, Grh, SR)
        y = y + .PixelHeight
    End With
    
    Grh = HeadData(UserHead).Head(Dir).GrhIndex
    
    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .PixelWidth
        SR.Bottom = SR.Top + .PixelHeight
        
        x = picPJ.Width * 0.5 - .PixelWidth * 0.5 + 4
        y = y + BodyData(UserBody).HeadOffset.y - .PixelHeight
        
        Call DrawTransparentGrhtoHdc(picPJ.hdc, x, y, Grh, SR)
    End With
End Sub

Private Sub DrawHead(ByVal Head As Integer, ByVal PicIndex As Integer)

    Dim SR As RECT
    Dim Grh As Long
    Dim x As Long
    Dim y As Long
    
    Call DrawImageInPicture(picHead(PicIndex), Picture, 0, 0, , , picHead(PicIndex).Left, picHead(PicIndex).Top)
    
    Grh = HeadData(Head).Head(Dir).GrhIndex

    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .PixelWidth
        SR.Bottom = SR.Top + .PixelHeight
        
        x = picHead(PicIndex).Width * 0.5 - .PixelWidth * 0.5 + 2.5
        y = 1
        
        Call DrawTransparentGrhtoHdc(picHead(PicIndex).hdc, x, y, Grh, SR)
    End With
    
End Sub

Private Sub NameTxt_GotFocus()
    NameTxtBorder.BorderColor = &HC0FFFF
    NameTxtBorder.Visible = True
End Sub

Private Sub NameTxt_LostFocus()
    NameTxtBorder.Visible = False
End Sub

Private Sub txtMail_GotFocus()
    txtMailBorder.BorderColor = &HC0FFFF
    txtMailBorder.Visible = True
End Sub

Private Sub txtMail_LostFocus()
    txtMailBorder.Visible = False
End Sub

Private Sub txtPasswd_GotFocus()
    txtPasswdBorder.BorderColor = &HC0FFFF
    txtPasswdBorder.Visible = True
End Sub

Private Sub txtPasswd_LostFocus()
    txtPasswdBorder.Visible = False
End Sub

Private Sub txtConfirmPasswd_GotFocus()
    txtConfirmPasswdBorder.BorderColor = &HC0FFFF
    txtConfirmPasswdBorder.Visible = True
End Sub

Private Sub txtConfirmPasswd_LostFocus()
    txtConfirmPasswdBorder.Visible = False
End Sub
