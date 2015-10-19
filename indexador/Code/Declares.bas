Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

'Objetos públicos


Public SurfaceDB As clsSurfaceManager   'No va new porque es unainterfaz, el new se pone al decidir que clase de objeto es


'' The main timer of the game.

Public LOOPActual As Long
Public GRHActual As Integer
Public DataIndexActual As Integer
Public Const VERSION_ACTUAL As String = "1.06"
'Sonidos


' Head index of the casper. Used to know if a char is killed
Public Const CASPER_HEAD As Integer = 500

Public LastFound As Long
Public BMPBuscado As Long

Public LoadingNew As Boolean

'Musica
Public Const MIdi_Inicio As Byte = 6

Public RawServersList As String
Public SavePath As Byte
Public DibujarFondo As Boolean
Public ColorFondo As Long

Public GrHCambiando As Boolean
Public TempGrh As GrhData
Public tempDataIndex As BodyData

Public Type tColor
    r As Byte
    G As Byte
    b As Byte
End Type

Type IndexacionActual
    Total As Integer
    Inicios(1 To 70) As Position
    activo As Boolean
    Ancho As Integer
    Alto As Integer
End Type

Public ColoresPJ(0 To 50) As tColor
Public CarpetaDeInis As String
Public CarpetaGraficos As String
Public DibujarIndexaciones As IndexacionActual

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
End Type

Public currentMidi As Long

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte

Public Trabajando As Boolean





Public Tips() As String * 255
Public Const LoopAdEternum As Integer = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum



Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Public Type ArchivoBMP
    BMPFileHeader As BITMAPINFOHEADER
    bmpInfo As BITMAPINFO
    BMPData() As Byte
End Type

Public Type ResoGrap
    offset As Long
    Archivo As Byte
    tamaño As Long
End Type
Public Type RecursoGrafico
    graficos() As ResoGrap
    UltimoGrafico As Long
End Type

Global Const DIB_RGB_COLORS = 0
Global Const MAXGrH = 32000

Public Nombres As Boolean

Public MixedKey As Long


Public DibujarWalk As Integer
'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As String
Public UserSexo As String
Public UserRaza As String
Public UserEmail As String



Public Musica As Boolean
Public Sound As Boolean

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean

Public UsingSkill As Integer
Public Enum e_EstadoIndexador
    Grh = 0
    Body = 1
    Cabezas = 2
    Cascos = 3
    Escudos = 4
    Armas = 5
    Botas = 6
    Capas = 7
    Fx = 8
    Resource = 9
End Enum
Public EstadoIndexador As e_EstadoIndexador

Public UltimoindexE(e_EstadoIndexador.Grh To e_EstadoIndexador.Resource) As Long
Public EstadoNoGuardado(e_EstadoIndexador.Grh To e_EstadoIndexador.Resource) As Boolean
Public MD5HushYo As String * 16




   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
End Enum
Public Enum EGuildPermiso
    VerBodeda = 1
    DepositarBodeda = 2
    RetirarBoveda = 3
    VerMiembro = 4
    AceptarMiembro = 5
    ExpulsarMiembro = 6
    CambiarGuildNews = 7
End Enum

Public cabezaActual As Integer
Public cuerpoActual As Integer

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'String contants
Public Const ENDC As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String
Public ResourceF As RecursoGrafico
Public ResourceFile As Byte
Public UsarGrhLong As Boolean
Public IniciadoTodo As Boolean
'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal DX As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceCuerpoLong
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Public Type tIndiceFxLong
    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer
End Type
