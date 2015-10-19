VERSION 5.00
Begin VB.Form frmMapOp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Propiedades del Mapa"
   ClientHeight    =   2340
   ClientLeft      =   5505
   ClientTop       =   4410
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2040
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox restringuir 
      Height          =   1230
      ItemData        =   "frmMapOp.frx":0000
      Left            =   2880
      List            =   "frmMapOp.frx":0013
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.ListBox terreno 
      Height          =   1230
      ItemData        =   "frmMapOp.frx":003C
      Left            =   120
      List            =   "frmMapOp.frx":0055
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtMusicNum 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Magia 
      Caption         =   "Magia sin efecto"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CheckBox PK 
      Caption         =   "Insegura"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox INVI 
      Caption         =   "Invi sin efecto"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox RESU 
      Caption         =   "Resu sin efecto"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtNameMap 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Restringir para:"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Terreno:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Número de MP3:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre del Mapa:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmMapOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            Audio.StopMidi
            Audio.PlayMIDI DirMidi & txtMusicNum.Text & ".mid"
        Case 1
            Audio.StopMidi
    End Select
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If MapInfo.PK = True Then
        PK.value = vbChecked
    Else
        PK.value = vbUnchecked
    End If
    
    If MapInfo.ResuSinEfecto = 1 Then
        RESU.value = vbChecked
    Else
        RESU.value = vbUnchecked
    End If
    
    If MapInfo.InviSinEfecto = 1 Then
        INVI.value = vbChecked
    Else
        INVI.value = vbUnchecked
    End If
    
    If MapInfo.MagiaSinEfecto = 1 Then
        Magia.value = vbChecked
    Else
        Magia.value = vbUnchecked
    End If

    txtNameMap.Text = MapInfo.name
    txtMusicNum.Text = MapInfo.Music
    
    terreno.ListIndex = MapInfo.terreno
    restringuir.ListIndex = MapInfo.restringir
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
