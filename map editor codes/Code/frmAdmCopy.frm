VERSION 5.00
Begin VB.Form frmAdmCopy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración de copiado"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRender 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   120
      ScaleHeight     =   345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   4
      Top             =   2040
      Width           =   6000
   End
   Begin VB.CommandButton cmdCopyList 
      Caption         =   "Copiar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   6015
   End
   Begin VB.CommandButton cmdSaveCopy 
      Caption         =   "Guardar Copiado"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copiar seleccion"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.ListBox lstCopy 
      Height          =   1230
      ItemData        =   "frmAdmCopy.frx":0000
      Left            =   120
      List            =   "frmAdmCopy.frx":0002
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
End
Attribute VB_Name = "frmAdmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopyList_Click()
    cData = cp
End Sub

Private Sub cmdSaveCopy_Click()
    Dim tmpStr As String
    Dim F As Integer
    
    tmpStr = InputBox("De un nombre para guardar el copiado")
    If tmpStr = "" Then Exit Sub
    
    F = FreeFile()
    Open App.Path & "\copys\" & tmpStr & ".cbm" For Binary Access Write As #F
        Put #F, , cData
    Close #F
    
    lstCopy.AddItem tmpStr & ".cbm"
End Sub

Private Sub Form_Load()
    Dim file As String
    
    file = Dir(App.Path & "\copys\*.cbm")
    
    While file <> ""
        lstCopy.AddItem file
        file = Dir
    Wend
End Sub

Private Sub lstCopy_Click()
    Dim F As Integer
    F = FreeFile()
    Open App.Path & "\copys\" & lstCopy.List(lstCopy.ListIndex) For Binary Access Read As #F
        Get #F, , cp.dX
        Get #F, , cp.dY
        
        ReDim cp.copied(cp.dX, cp.dY)
        
        Seek #F, 1
        Get #F, , cp
    Close #F
    
    Dim DestRect As RECT
    
    DestRect.bottom = 345
    DestRect.Right = 400
    DestRect.Left = 0
    DestRect.Top = 0
    
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        Engine.RenderCopy
    D3DDevice.EndScene
    D3DDevice.Present DestRect, ByVal 0, picRender.hwnd, ByVal 0
    
    cmdCopyList.Enabled = True
End Sub
