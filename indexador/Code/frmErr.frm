VERSION 5.00
Begin VB.Form frmErr 
   Caption         =   "Buscar Errores de Indexacion"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long
Dim MsgErr As String
Dim ErrGrh As ErroresGrh
Me.Visible = True
For i = 1 To MAXGrH
    If GrhData(i).NumFrames <> 0 Then
        Call GrhCorrecto(GrhData(i), MsgErr, ErrGrh)
        If MsgErr <> "" Then
            List1.AddItem i & "(" & MsgErr & ")"
        End If
    End If
    DoEvents
Next i

End Sub

Private Sub List1_Click()
frmMain.Lista.listIndex = Val(List1.List(List1.listIndex))
End Sub
