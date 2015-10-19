VERSION 5.00
Begin VB.Form frmUnion 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Union en mapas adyacentes"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   120
      Picture         =   "frmUnion.frx":0000
      ScaleHeight     =   3900
      ScaleWidth      =   5820
      TabIndex        =   18
      Top             =   3120
      Width           =   5820
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "APL ICAR"
      Height          =   1695
      Left            =   5640
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "CANCELAR"
      Height          =   1695
      Left            =   5280
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   10
      Text            =   "0"
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   8
      Text            =   "0"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   7
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Text            =   "0"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CheckBox Aplicar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox AutoMapeo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   120
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   1320
      Width           =   495
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      X1              =   288
      X2              =   288
      Y1              =   200
      Y2              =   8
   End
End
Attribute VB_Name = "frmUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PDsUp As Integer = 90
Private Const PDsDown As Integer = 11
Private Const PDsLeft As Integer = 89
Private Const PDsRight As Integer = 12

Private Const PScUp As Integer = 10
Private Const PScDown As Integer = 91
Private Const PScLeft As Integer = 90
Private Const PScRight As Integer = 11

Private Sub cmdAplicar_Click()
Dim y As Long, x As Long, x1 As Integer, y1 As Integer

' ARRIBA
If Mapa(0).Text > -1 And Aplicar(0).value = 1 Then
    If AutoMapeo(0).value = vbChecked Then
        Map_Load_In2 frmMain.FileMapDir & "mapa" & Mapa(0).Text & ".abr"
        
        For x = 1 To 100
            y1 = PDsUp - 9
            For y = 1 To PScUp
                MapData(x, y) = MapData2(x, y1)
                
                y1 = y1 + 1
            Next y
        Next x
    End If
    
    y = PScUp
    For x = (PScRight + 1) To (PScLeft - 1)
        If MapData(x, y).Blocked = 0 Then
            MapData(x, y).TileExit.map = Mapa(0).Text
            If Mapa(0).Text = 0 Then
                MapData(x, y).TileExit.x = 0
                MapData(x, y).TileExit.y = 0
            Else
                MapData(x, y).TileExit.x = x
                MapData(x, y).TileExit.y = PDsUp
            End If
        End If
    Next
End If

' DERECHA
If Val(Mapa(1).Text) > -1 And Aplicar(1).value = 1 Then
    If AutoMapeo(1).value = vbChecked Then
        Map_Load_In2 frmMain.FileMapDir & "mapa" & Mapa(1).Text & ".abr"
        For y = 1 To 100
            x1 = PDsRight
            For x = PScLeft To PScLeft + 10
                MapData(x, y) = MapData2(x1, y)
                
                x1 = x1 + 1
            Next x
        Next y
    End If
    
    x = PScLeft
    For y = (PScUp + 1) To (PScDown - 1)
        If MapData(x, y).Blocked = 0 Then
            MapData(x, y).TileExit.map = Mapa(1).Text
                If Mapa(1).Text = 0 Then
                    MapData(x, y).TileExit.x = 0
                    MapData(x, y).TileExit.y = 0
                Else
                    MapData(x, y).TileExit.x = PDsRight
                    MapData(x, y).TileExit.y = y
                End If
        End If
    Next
End If

' ABAJO
If Mapa(2).Text > -1 And Aplicar(2).value = 1 Then
    Map_Load_In2 frmMain.FileMapDir & "mapa" & Mapa(2).Text & ".abr"
    If AutoMapeo(2).value = vbChecked Then
        For x = 1 To 100
            y1 = PDsDown
            For y = PScDown To PScDown + 9
                MapData(x, y) = MapData2(x, y1)
                
                y1 = y1 + 1
            Next y
        Next x
    End If
    
    y = PScDown
    For x = (PScRight + 1) To (PScLeft - 1)
        If MapData(x, y).Blocked = 0 And MapData2(x, PDsDown).Blocked = 0 Then
            MapData(x, y).TileExit.map = Mapa(2).Text
            If Mapa(2).Text = 0 Then
                MapData(x, y).TileExit.x = 0
                MapData(x, y).TileExit.y = 0
            Else
                MapData(x, y).TileExit.x = x
                MapData(x, y).TileExit.y = PDsDown
            End If
        End If
    Next
End If

' IZQUIERDA
If Mapa(3).Text > -1 And Aplicar(3).value = 1 Then
    If AutoMapeo(3).value = vbChecked Then
        Map_Load_In2 frmMain.FileMapDir & "mapa" & Mapa(3).Text & ".abr"
        y1 = PScLeft + 10
        For y = 1 To 100
            x1 = PDsLeft - 10
            For x = 1 To 11
                MapData(x, y) = MapData2(x1, y)
                
                x1 = x1 + 1
            Next x
        Next y
    End If
    
    x = PScRight
    For y = (PScUp + 1) To (PScDown - 1)
        If MapData(x, y).Blocked = 0 Then
            MapData(x, y).TileExit.map = Mapa(3).Text
            If Mapa(3).Text = 0 Then
                MapData(x, y).TileExit.x = 0
                MapData(x, y).TileExit.y = 0
            Else
                MapData(x, y).TileExit.x = PDsLeft
                MapData(x, y).TileExit.y = y
            End If
        End If
    Next
End If
End Sub


Private Sub Form_Load()
    Me.Top = 0
    Me.Left = Screen.Width - Me.Width
End Sub

