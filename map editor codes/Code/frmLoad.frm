VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Editor • Cargando"
   ClientHeight    =   345
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   0
      Top             =   0
      Width           =   5390
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cargando"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -15
         TabIndex        =   1
         Top             =   30
         Width           =   5415
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   240
         Left            =   30
         Top             =   30
         Width           =   5250
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
