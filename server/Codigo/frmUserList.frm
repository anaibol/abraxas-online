VERSION 5.00
Begin VB.Form frmUserList 
   BackColor       =   &H00404040&
   ClientHeight    =   4665
   ClientLeft      =   3600
   ClientTop       =   2190
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5520
   Begin VB.CommandButton Command2 
      Caption         =   "Echar todos los no Logged"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualiza"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim LoopC As Integer
    
    Text2.Text = "MaxPoblacion: " & MaxPoblacion & vbNewLine
    Text2.Text = Text2.Text & "LastUser: " & LastUser & vbNewLine
    Text2.Text = Text2.Text & "Poblacion: " & Poblacion & vbNewLine
    
    List1.Clear
    
    For LoopC = 1 To MaxPoblacion
        List1.AddItem Format(LoopC, "000") & " " & IIf(UserList(LoopC).flags.Logged, UserList(LoopC).Name, vbNullString)
        List1.ItemData(List1.NewIndex) = LoopC
    Next LoopC
End Sub

Private Sub Command2_Click()
    Dim LoopC As Integer
    
    For LoopC = 1 To MaxPoblacion
        If UserList(LoopC).ConnID <> -1 And Not UserList(LoopC).flags.Logged Then
            Call CloseSocket(LoopC)
        End If
    Next LoopC
End Sub

Private Sub List1_Click()
    Dim UserIndex As Integer
    If List1.ListIndex <> -1 Then
        UserIndex = List1.ItemData(List1.ListIndex)
        If UserIndex > 0 And UserIndex <= MaxPoblacion Then
            With UserList(UserIndex)
                Text1.Text = "Logged: " & .flags.Logged & vbNewLine
                Text1.Text = Text1.Text & "IdleCount: " & .Counters.IdleCount & vbNewLine
                Text1.Text = Text1.Text & "ConnId: " & .ConnID & vbNewLine
                Text1.Text = Text1.Text & "ConnIDValida: " & .ConnIDValida & vbNewLine
            End With
        End If
    End If
End Sub

