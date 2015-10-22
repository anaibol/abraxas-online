VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ButtonHolder As Image
Attribute ButtonHolder.VB_VarHelpID = -1

Private RollOverButtonImage As Picture  'img holder
Private NormalButtonImage As Picture 'img holder
Private ClickButtonImage As Picture  'img holder
Private DisabledButtonImage As Picture  'img holder

Private SoundClick As Boolean

Private FormActual As Form 'Form owner

Private OverButton As Boolean 'Indicates wether the mouse is over the button or not
Private Disabled As Boolean 'Indicates wether it's enabled or not

Public Sub Initialize(ByRef imgButton As Image, ByVal imgNormalPath As String, _
            ByVal imgRollOverPath As String, ByVal imgClickPath As String, ByRef Formulario As Form, _
            Optional ByVal imgDisabledPath As String, Optional ByVal StartDisabled As Boolean = False, _
            Optional ByVal MakeClickSound As Boolean = True, Optional ByVal ShowHandPointer As Boolean = True)
'Sets the reference to the image container

    Set ButtonHolder = imgButton
    Set FormActual = Formulario
    
    Set RollOverButtonImage = LoadPicture(imgRollOverPath)
    Set NormalButtonImage = LoadPicture(imgNormalPath)
    Set ClickButtonImage = LoadPicture(imgClickPath)
    Set DisabledButtonImage = LoadPicture(imgDisabledPath)
    
    SoundClick = MakeClickSound
    
    With ButtonHolder
        If ShowHandPointer Then
            .MousePointer = vbCustom
            .MouseIcon = picMouseIcon
        End If
        
        If StartDisabled Then
            Disabled = True
            .Picture = DisabledButtonImage
        Else
            .Picture = NormalButtonImage
        End If
    End With
    
End Sub

Private Sub ButtonHolder_Click()
'Implements the mouse click event of the image button.

'Plays the Click Sound.
    If Disabled Then
        Exit Sub
    End If
    
    If Not SoundClick Then
        Exit Sub
    End If
    
    Call Audio.Play(SND_CLICK)
End Sub

Private Sub ButtonHolder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Implements the mouse down event of the image button.
'Toggle to Click image button.
    If Disabled Then
        Exit Sub
    End If
    ButtonHolder.Picture = ClickButtonImage
End Sub

Private Sub ButtonHolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Implements the mouse move event of the image button.
'Toggle to pressed image button.

    If Disabled Then
        Exit Sub
    End If
    
    If Not OverButton Then
    
        'This helps when two buttons are close
        FormActual.LastPressed.ToggleToNormal
        
        ButtonHolder.Picture = RollOverButtonImage
        OverButton = True
        
        Set FormActual.LastPressed = Me
    End If
End Sub

Public Sub ToggleToNormal()
'Toggle to normal image button.

    If Disabled Then
        Exit Sub
    End If
    
    If Not OverButton Then
        Exit Sub
    End If
    
    ButtonHolder.Picture = NormalButtonImage
    OverButton = False
End Sub

Private Sub ButtonHolder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Implements the mouse up event of the image button.
'Toggle to RollOver image button.

    If Disabled Then
        Exit Sub
    End If
    
    ButtonHolder.Picture = RollOverButtonImage
End Sub

Public Sub EnableButton(Optional ByVal ToogleEnable As Boolean = True)
    If Disabled = Not ToogleEnable Then
        Exit Sub
    End If
    
    OverButton = False
    
    Disabled = Not ToogleEnable
    
    If Disabled Then
        ButtonHolder.Picture = DisabledButtonImage
    Else
        ButtonHolder.Picture = NormalButtonImage
    End If
End Sub

Public Function IsEnabled() As Boolean
    IsEnabled = Not Disabled
End Function
