VERSION 5.00
Begin VB.Form frmSalir 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image cancelar 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      MouseIcon       =   "frmSalir.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image aceptar 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmSalir.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "frmSalir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Aceptar_Click()
    Call CloseClient
End Sub

Private Sub Cancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(GrhPath & "Salir.jpg")
Call Make_Transparent_Form(Me.hWnd, 200)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hWnd)
Else
    Unload Me
End If
End Sub
