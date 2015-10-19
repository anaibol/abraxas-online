VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion de Directorios"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdComenzar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton cmdDat 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtDat 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtInit 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Directorios de dats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Directorios de cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1845
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdComenzar_Click()
    If txtDat.Text = "" Then
        MsgBox "Directorio de Dats inválido"
        Exit Sub
    End If
    
    If txtInit.Text = "" Then
        MsgBox "Directorio de Inits inválido"
        Exit Sub
    End If
    
    If Right$(txtInit.Text, 1) <> "\" Then
        txtInit.Text = txtInit.Text & "\"
    End If
    
    If Right$(txtDat.Text, 1) <> "\" Then
        txtDat.Text = txtDat.Text & "\"
    End If
    
    Call General_Var_Write(App.Path & "\MapEditor.ini", "CONFIG", "Cliente", txtInit.Text)
    Call General_Var_Write(App.Path & "\MapEditor.ini", "CONFIG", "Dats", txtDat.Text)

    Unload Me

    Call Main
End Sub

Private Sub cmdDat_Click()
    txtDat.Text = OpenFolderSearch()
End Sub

Private Sub cmdInit_Click()
    txtInit.Text = OpenFolderSearch()
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = Screen.Width - Me.Width
End Sub
