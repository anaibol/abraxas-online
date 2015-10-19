VERSION 5.00
Begin VB.Form frmMinimap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Minimap"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMinimap 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   0
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   0
      Top             =   0
      Width           =   1530
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   135
         Left            =   1560
         TabIndex        =   1
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMinimap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub tmrMinimap_Timer()
    DibujarMiniMapa
End Sub
