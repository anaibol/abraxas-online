VERSION 5.00
Begin VB.Form frmAuto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Automatizadores"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fCaminos 
      Caption         =   "Caminos"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      Begin VB.CommandButton cmdInsertCam 
         Caption         =   "Insertar Camino"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   3720
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdInsertCam_Click()

End Sub

Function Insertar_Superficie(ByVal i As Integer, ByVal layer As Byte, ByVal destX As Integer, ByVal destY As Integer) As Boolean
    Dim tX As Integer, tY As Integer, despTile As Integer
        
    For tY = destY To destY + SupData(i).Height
        For tX = destX To destX + SupData(i).Width
            MapData(tX, tY).Graphic(layer).GrhIndex = CInt(Val(SupData(i).Grh) + despTile)
             
            despTile = despTile + 1
        Next x
    Next y
End Function
