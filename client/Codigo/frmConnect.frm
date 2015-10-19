VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConnect 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Abraxas"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   330
   ClientWidth     =   12000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   0
   End
   Begin VB.TextBox ServerIpe 
      Height          =   390
      Left            =   9840
      TabIndex        =   6
      Text            =   "127.0.0.1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   7080
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin RichTextLib.RichTextBox NameTxt 
      Height          =   360
      Left            =   4875
      TabIndex        =   1
      Top             =   4080
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   635
      _Version        =   393217
      BackColor       =   16576
      BorderStyle     =   0
      MultiLine       =   0   'False
      MousePointer    =   1
      DisableNoScroll =   -1  'True
      MaxLength       =   15
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmConnect.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox PasswordTxt 
      Height          =   345
      Left            =   4875
      TabIndex        =   0
      Top             =   4770
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   609
      _Version        =   393217
      BackColor       =   16576
      BorderStyle     =   0
      MultiLine       =   0   'False
      MousePointer    =   1
      DisableNoScroll =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmConnect.frx":0090
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Entrar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   5040
      TabIndex        =   10
      Top             =   5250
      Width           =   615
   End
   Begin VB.Label Empezar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Empezar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   6240
      TabIndex        =   9
      Top             =   5250
      Width           =   855
   End
   Begin VB.Label PsNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   4725
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label NmNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   4005
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape PasswordTxtBorder 
      BorderColor     =   &H0080FFFF&
      Height          =   315
      Left            =   4830
      Top             =   4800
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Shape NameTxtBorder 
      BorderColor     =   &H0080FFFF&
      Height          =   315
      Left            =   4830
      Top             =   4080
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Image SavePassImg 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   4320
      Picture         =   "frmConnect.frx":0114
      Top             =   5880
      Width           =   180
   End
   Begin VB.Image imgCrearPJ 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      MousePointer    =   99  'Custom
      Top             =   5250
      Width           =   1275
   End
   Begin VB.Image imgConectar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MousePointer    =   99  'Custom
      Top             =   5250
      Width           =   1125
   End
   Begin VB.Label ClickSavePass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   5640
      Width           =   255
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ret As Long

Private arrValues() As String

'Flag para la tecla BackSpace
Private bKeyBack As Boolean

Private eName As String
Private ePass As String

Private Sub ClickSavePass_Click()
    If Not SavePassImg.Visible Then
        SavePassImg.Visible = True
    End If
End Sub

Private Sub Command1_Click()
    EstadoLogin = Recuperando
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If

    frmMain.Socket1.Startup
            
    frmMain.Socket1.HostName = ServerIP
    frmMain.Socket1.RemotePort = ServerPort

    frmMain.Socket1.Connect
End Sub

Private Sub Empezar_Click()

    'If frmMain.Socket1.Connected Then
    '    frmMain.Socket1.Disconnect
    '    frmMain.Socket1.Cleanup
    '    DoEvents
    'End If
    
    'frmMain.Socket1.Startup
        
    'frmMain.Socket1.HostName = ServerIP
    'frmMain.Socket1.RemotePort = ServerPort
    
    'frmMain.Socket1.Connect
    
    Call Audio.mSound_PlayWav(SND_CLICK)
    
    frmCrearPersonaje.Show vbModal
End Sub

Private Sub Entrar_Click()

On Error Resume Next

    Visible = False
    
    Call Audio.mSound_PlayWav(SND_CLICK)
    
    If LenB(PasswordTxt.text) < 1 Then
        PasswordTxtBorder.BorderColor = &H80&
        PasswordTxtBorder.Visible = True
        PsNo.Visible = True
        PasswordTxt.SetFocus
        Exit Sub
    End If
       
    If LenB(NameTxt.text) < 3 Or Len(NameTxt.text) > 15 Then
        NameTxtBorder.BorderColor = &H80&
        NameTxtBorder.Visible = True
        NmNo.Visible = True
        NameTxt.SetFocus
        Exit Sub
    End If
    
    If Len(PasswordTxt.text) < 6 Then
        PasswordTxtBorder.BorderColor = &H80&
        PasswordTxtBorder.Visible = True
        PsNo.Visible = True
        PasswordTxt.SetFocus
        Exit Sub
    End If
    
    Timer2.Enabled = False
    
    PasswordTxt.Visible = False
    NameTxt.Visible = False
    MousePointer = 11
            
    'Update user info
    UserName = NameTxt.text

    UserPassword = PasswordTxt.text
        
    EstadoLogin = Normal
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If

    frmMain.Socket1.Startup
            
    frmMain.Socket1.HostName = ServerIP
    frmMain.Socket1.RemotePort = ServerPort

    frmMain.Socket1.Connect
End Sub

Private Sub Form_Activate()
    If Not ChangeResolution Then
        ret = GetWindowLong(hwnd, -20)
        ret = ret Or &H80000
        SetWindowLong hwnd, -20, ret
        Timer1.Interval = 5
        Timer1.Enabled = True
    End If
    
    Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        prgRun = False
    End If
End Sub

Private Sub Form_Load()

On Error Resume Next

    Picture = LoadPicture(GrhPath & "Conectar.jpg")
     
    Icon = frmMain.Icon
    
    Call Make_Transparent_Richtext(NameTxt.hwnd)
    Call Make_Transparent_Richtext(PasswordTxt.hwnd)
    
    'Get the username/password
    eName = (GetVar(DataPath & "Game.ini", "INIT", "Name"))
    
    If LenB(eName) > 0 Then
        NameTxt.text = eName
        
        ePass = (GetVar(DataPath & "Game.ini", "INIT", "Pass"))
            
        If LenB(ePass) > 0 Then
            PasswordTxt.text = ePass
            SavePassImg.Visible = True
            Call SendMessage(NameTxt.hwnd, &H7, ByVal 0&, ByVal 0&)
        Else
            Call SendMessage(PasswordTxt.hwnd, &H7, ByVal 0&, ByVal 0&)
            'PasswordTxtBorder.Visible = True
        End If
        
    Else
        NameTxt.SelLength = Len(NameTxt.text)
        NameTxt.SelColor = RGB(255, 255, 255)
               
        PasswordTxt.SelLength = Len(PasswordTxt.text)
        PasswordTxt.SelColor = RGB(255, 255, 255)
        Call SendMessage(NameTxt.hwnd, &H7, ByVal 0&, ByVal 0)
    End If
    
    'Redim arrValues(0)
    'cargar los valores desde el archivo de texto en el array
    'Call LoadValues(arrValues)
End Sub

Private Sub imgConectar_Click()
    Call Audio.mSound_PlayWav(SND_CLICK)
    
    If LenB(PasswordTxt.text) < 1 Then
        PasswordTxtBorder.BorderColor = &H80&
        PasswordTxtBorder.Visible = True
        PsNo.Visible = True
        PasswordTxt.SetFocus
        Exit Sub
    End If
       
    If LenB(NameTxt.text) < 3 Or Len(NameTxt.text) > 15 Then
        NameTxtBorder.BorderColor = &H80&
        NameTxtBorder.Visible = True
        NmNo.Visible = True
        NameTxt.SetFocus
        Exit Sub
    End If
    
    If Len(PasswordTxt.text) < 6 Then
        PasswordTxtBorder.BorderColor = &H80&
        PasswordTxtBorder.Visible = True
        PsNo.Visible = True
        PasswordTxt.SetFocus
        Exit Sub
    End If
    
    Timer2.Enabled = False
    
    PasswordTxt.Visible = False
    NameTxt.Visible = False
    MousePointer = 11
            
    'Update user info
    UserName = NameTxt.text

    UserPassword = PasswordTxt.text
        
    EstadoLogin = Normal
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If

    frmMain.Socket1.Startup
            
    frmMain.Socket1.HostName = ServerIP
    frmMain.Socket1.RemotePort = ServerPort

    frmMain.Socket1.Connect
End Sub

Private Sub Empezar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Empezar.ForeColor = &H80FFFF
    Entrar.ForeColor = &HC0FFFF
End Sub

Private Sub Entrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Entrar.ForeColor = &H80FFFF
    Empezar.ForeColor = &HC0FFFF
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Empezar.ForeColor = &HC0FFFF
    Entrar.ForeColor = &HC0FFFF
End Sub

Private Sub imgCrearPJ_Click()
    Call Audio.mSound_PlayWav(SND_CLICK)
    
    frmCrearPersonaje.Show vbModal
End Sub

Private Sub NameTxt_Change()

    Dim i As Long
    Dim tempstr As String
    Dim CharAscii As Integer
    
    If LenB(NameTxt.text) > 0 Then
    
        For i = 1 To Len(NameTxt.text)
            CharAscii = Asc(mid$(NameTxt.text, i, 1))
            
            If (CharAscii > 64 And CharAscii < 91) Or (CharAscii > 96 And CharAscii < 123) Or CharAscii = 32 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        tempstr = StrConv(tempstr, vbProperCase)
        
        If LenB(tempstr) < 1 Then
            NameTxt.text = vbNullString
            Exit Sub
        End If
        
        If tempstr <> NameTxt.text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            NameTxt.text = tempstr
            PasswordTxtBorder.BorderColor = &HC0FFFF
        End If

        If LenB(eName) > 0 Then
            If LenB(ePass) > 0 Then
                If NameTxt.text = eName Then
                    PasswordTxt.text = ePass
                    SavePassImg.Visible = True
                Else
                    PasswordTxt.text = vbNullString
                    PasswordTxt.SelColor = RGB(255, 255, 255)
                    SavePassImg.Visible = False
                End If
            Else
                PasswordTxt.text = vbNullString
                PasswordTxt.SelColor = RGB(255, 255, 255)
                SavePassImg.Visible = False
            End If
        Else
            SavePassImg.Visible = False
            PasswordTxt.text = vbNullString
            PasswordTxt.SelColor = RGB(255, 255, 255)
        End If
    Else
        SavePassImg.Visible = False
        PasswordTxt.text = vbNullString
        PasswordTxt.SelColor = RGB(255, 255, 255)
    End If
        
    'If LenB(NameTxt) < 1 Then
    'Call AutoCompletar_TextBox
    'End If
    
    NameTxt.SelLength = Len(NameTxt.text)
        
    NameTxt.SelColor = RGB(255, 255, 255)
    
    NameTxt.SelStart = Len(NameTxt.text)
    
    NmNo.Visible = False
    NameTxtBorder.BorderColor = &HC0FFFF
    
    Timer2.Enabled = False
    Timer2.Interval = 0
    Timer2.Interval = 500
    Timer2.Enabled = True
End Sub

Private Sub PasswordTxt_Change()
    
    SavePassImg.Visible = False
    
    PasswordTxt.SelStart = 0

    PasswordTxt.SelLength = Len(PasswordTxt.text)
        
    PasswordTxt.SelColor = RGB(255, 255, 255)
    
    PasswordTxt.SelStart = Len(PasswordTxt.text)
    
    PsNo.Visible = False
    
    PasswordTxtBorder.BorderColor = &HC0FFFF
    
    Call SendMessage(PasswordTxt.hwnd, &HCC, Asc("*"), 0)
    
    HideCaret PasswordTxt.hwnd
End Sub

Private Sub NameTxt_GotFocus()
    NameTxtBorder.Visible = True
    
    If NameTxtBorder.BorderColor = &H80& Then
        NmNo.Visible = True
        Exit Sub
    End If
        
    NameTxt.SelStart = 0
    NameTxt.SelLength = Len(NameTxt.text)
End Sub

Private Sub PasswordTxt_GotFocus()

    If LenB(NameTxt.text) < 1 Then
        NameTxtBorder.BorderColor = &H80&
        NameTxtBorder.Visible = True
        NmNo.Visible = True
        NameTxt.SetFocus
    End If
    
    PasswordTxtBorder.Visible = True
    HideCaret PasswordTxt.hwnd

    If PasswordTxtBorder.BorderColor = &H80& Then
        PsNo.Visible = True
    End If
End Sub

Private Sub NameTxt_LostFocus()
    NameTxtBorder.Visible = False
    NmNo.Visible = False
    PsNo.Visible = False
    PasswordTxtBorder.BorderColor = &HC0FFFF
End Sub

Private Sub PasswordTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    HideCaret PasswordTxt.hwnd
End Sub

Private Sub PasswordTxt_LostFocus()
    PasswordTxtBorder.Visible = False
    PsNo.Visible = False
End Sub

Private Sub NameTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call imgConectar_Click
    End If
End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call imgConectar_Click
    ElseIf KeyCode = vbKeyBack Then
        PasswordTxt.text = vbNullString
    End If
    
    HideCaret PasswordTxt.hwnd
End Sub

Private Sub PasswordTxt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HideCaret PasswordTxt.hwnd
End Sub

Private Sub SavePassImg_Click()
    SavePassImg.Visible = Not SavePassImg.Visible
End Sub

'Rutina para cargar los valores desde el archivo .dat
Private Sub LoadValues(pArrValues() As String)
    Dim lIndex As Long
    Dim sValue As String
    Dim nfile As Integer
    nfile = FreeFile()

    If Len(Dir(DataPath & "Nombres.ini")) = 0 Then
        Exit Sub
    End If
   
    'Leemos los nombres desde el archivo de datos.
    Open DataPath & "Nombres.ini" For Input As nfile

    'Leer las lineas del archivo
    While Not EOF(1)
        Line Input #1, sValue
        'Agregar nuevo valor
        ReDim Preserve pArrValues(lIndex)
        If sValue <> vbNullString Then
            pArrValues(lIndex) = sValue
        End If
        lIndex = lIndex + 1
    Wend
    Close
End Sub

Private Sub AutoCompletar_TextBox()
 
    Dim i As Integer
    Dim posSelect As Integer
 
    Select Case (bKeyBack Or Len(NameTxt.text) = 0)
        Case True
            bKeyBack = False
            Exit Sub
    End Select
    
    With NameTxt
        'Recorremos todos los elementos del array
        For i = 0 To UBound(arrValues)
            'Buscamos coincidencias
            If InStr(1, arrValues(i), .text, vbTextCompare) = 1 Then
                posSelect = .SelStart
                'Asignar el texto de array al textbox
                .text = arrValues(i)
           
                'seleccionar el texto
                .SelStart = posSelect
                .SelLength = Len(.text) - posSelect
                Exit For 'salimos del bucle
            End If
        Next i
    End With
End Sub

'Rutina para guardar los valores en el archivo de historial
Private Sub saveValues()
    Dim lIndex As Long
    Dim nfile As Integer
    nfile = FreeFile()

    'Redimensionar y preservar el array para añadir el nuevo valor
    lIndex = UBound(arrValues) + 1
    ReDim Preserve arrValues(lIndex)
   
    arrValues(lIndex) = NameTxt
   
    'Abrir el archivo para guardar los datos
    Open DataPath & "Nombres.ini" For Output As nfile
   
    Dim i As Integer
   
    'Recorrer le vector
    For i = 0 To UBound(arrValues)
        If LenB(arrValues(i)) > 0 Then
            Print nfile, arrValues(i)
        End If
    Next
    Close
End Sub

Private Sub DeleteValues()
   
    Dim lIndex As Long
    Dim sValue As String
   
    Dim sPath As String
   
    Kill App.path & "\Nombres.ini"
    Close
    
    'Esto es para ponerlo en un boton o checkboxz para borrar con este PUBLIC SUB
    Call DeleteValues
    NameTxt.text = vbNullString
    ReDim arrValues(0)
    Call LoadValues(arrValues)

End Sub

Private Sub Timer1_Timer()
    Static Cont As Integer

    Cont = Cont + 15
    If Cont > 255 Then
        Cont = 0
        Timer1.Enabled = False
    Else
        If Cont + 80 > 255 Then
            SetLayeredWindowAttributes hwnd, 0, 255, &H2
        Else
            SetLayeredWindowAttributes hwnd, 0, Cont + 80, &H2
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    If Not Screen.ActiveControl Is Nothing Then
        If Screen.ActiveControl = NameTxt Then
              HideCaret NameTxt.hwnd
        ElseIf Screen.ActiveControl = PasswordTxt Then
            HideCaret PasswordTxt.hwnd
        Else
            If LenB(NameTxt.text) > 0 Then
                NameTxt.text = eName
        
                If LenB(PasswordTxt.text) > 0 Then
                    NameTxt.SelStart = 0
                    NameTxt.SelLength = Len(NameTxt.text)
                    Call SendMessage(NameTxt.hwnd, &H7, ByVal 0&, ByVal 0&)
                Else
                    PasswordTxt.SelStart = 0
                    PasswordTxt.SelLength = Len(PasswordTxt.text)
                    Call SendMessage(PasswordTxt.hwnd, &H7, ByVal 0&, ByVal 0&)
                End If
                
            Else
                Call SendMessage(NameTxt.hwnd, &H7, ByVal 0&, ByVal 0)
            End If
        End If
    Else
        If LenB(NameTxt.text) > 0 Then
            NameTxt.text = eName
    
            If LenB(PasswordTxt.text) > 0 Then
                NameTxt.SelStart = 0
                NameTxt.SelLength = Len(NameTxt.text)
                Call SendMessage(NameTxt.hwnd, &H7, ByVal 0&, ByVal 0&)
            Else
                PasswordTxt.SelStart = 0
                PasswordTxt.SelLength = Len(PasswordTxt.text)
                Call SendMessage(PasswordTxt.hwnd, &H7, ByVal 0&, ByVal 0&)
            End If
            
        Else
            Call SendMessage(NameTxt.hwnd, &H7, ByVal 0&, ByVal 0)
        End If
    End If
    
    
End Sub
