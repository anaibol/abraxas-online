Attribute VB_Name = "Encryptions"
Option Explicit

'These are only used if the PacketEncType is not PacketEncTypeNone
Private Const PacketEncKey1 As String = "If the doors of perception were cleanSed," 'First encryption key (any string works)
Private Const PacketEncKey2 As String = " every thing would appear to man as it is: infInite." 'Second encryption key (any string works)
Public Const PacketEncSeed As Long = -47954995    'The number to start from (any random value works)
Public Const PacketEncKeys As Byte = 50     'Number of packet encryption keys

'***** RC4 *****
Private m_sBoxRC4(0 To 255) As Integer

'***** MISC *****

'Key-dependant
Private m_KeyS As String

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)

Public Sub GenerateEncryptionKeys(ByRef PacketKeys() As String)
'Generates a series of unique keys baSed off the parameters
'It is recommended you change this routine a bit for better safety for public games
'Do NOT use random (Rnd) values since the server and client must make identical keys

    Dim Seed As Long
    Dim Key1 As String
    Dim Key2 As String
    Dim B2() As Byte
    Dim b() As Byte
    Dim i As Long
    Dim j As Long

    'Set the start values
    Seed = PacketEncSeed
    Key1 = PacketEncKey1
    Key2 = PacketEncKey2
    
    'Set the number of keys
    ReDim PacketKeys(0 To PacketEncKeys - 1)
    
    'Crop down the keys if needed
    If Len(Key2) > 32 Then
        Key2 = Left$(Key2, 32)
    End If
    
    If Len(Key1) > 32 Then
        Key1 = Left$(Key1, 32)
    End If
    
    'Loop through the needed keys
    For i = 0 To PacketEncKeys - 1
    
        'Generate a new seed
        Seed = (i * Seed) - 1
    
        'Jumble up the keys through XOR randomization
        b = StrConv(Key1, vbFromUnicode) 'Convert to a byte array
        B2 = StrConv(Key2, vbFromUnicode)
        For j = 0 To Len(Key1) - 1
            Seed = Seed + j + 1         'Modify the seed baSed on the placement of the Char
            Do While Seed > 255         'Keep it in the byte range
                Seed = Seed - 255
            Loop
            b(j) = b(j) Xor Seed        'XOR the Char by the seed
            B2(j) = B2(j) Xor CByte(Seed * 0.5)
        Next j
        Key1 = StrConv(b, vbUnicode)     'Convert back to a string
        Key2 = StrConv(B2, vbUnicode)
            
        'Jumble up the keys through encryption
        Key2 = Encryption_RC4_EncryptString(Key2, Key1)
        Key1 = Encryption_RC4_EncryptString(Key1, Key2)
        
        'Store the key
        PacketKeys(i) = Key1
        
    Next i

End Sub

Public Function Encryption_RC4_DecryptString(Text As String, Optional key As String) As String
'Decrypts a string array with RC4 encryption

    Dim ByteArray() As Byte

    'Convert the data into a byte array
    ByteArray() = StrConv(Text, vbFromUnicode)

    'Decrypt the byte array
    Call Encryption_RC4_EncryptByte(ByteArray(), key)

    'Convert the byte array back into a string
    Encryption_RC4_DecryptString = StrConv(ByteArray(), vbUnicode)

End Function

Public Sub Encryption_RC4_EncryptByte(ByteArray() As Byte, Optional key As String)
'Encrypts a byte array with RC4 encryption

    Dim i As Long
    Dim j As Long
    Dim temp As Byte
    Dim Offset As Long
    Dim OrigLen As Long
    Dim sBox(0 To 255) As Integer

    'Set the new key (optional)
    If Len(key) > 0 Then
        Encryption_RC4_SetKey key
    End If
    
    'Create a local copy of the sboxes, this
    'is much more elegant than recreating
    'before encrypting/decrypting anything
    Call CopyMem(sBox(0), m_sBoxRC4(0), 512)

    'Get the size of the source array
    OrigLen = UBound(ByteArray) + 1

    'Encrypt the data
    For Offset = 0 To (OrigLen - 1)
        i = (i + 1) Mod 256
        j = (j + sBox(i)) Mod 256
        temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = temp
        ByteArray(Offset) = ByteArray(Offset) Xor (sBox((sBox(i) + sBox(j)) Mod 256))
    Next

End Sub

Public Function EncryptByte(ByVal value As Byte, Optional ByVal UserIndex As Integer = 0) As Byte
'SEGURIDAD
    Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, key() As Byte, ByteArray() As Byte, temp As Byte, Password As String
    
    If UserIndex > 0 Then
        With UserList(UserIndex)
            If .flags.Logged Then
                Password = mid$(mid$(.flags.Password, 5) & _
                    Right$(CStr(.Char.CharIndex * (-1) - .Pos.Map + .Char.CharIndex), 1) & _
                    Right$(.Name, 2) & _
                    Left$(CStr(.Char.CharIndex - .Pos.Map * 7 * Val(Right$(CStr(.Char.CharIndex), 1))), 1) & _
                    UCase$(Left$(.Name, 2)) & _
                    "49uN738D37H" _
                    , 3, 12)
            Else
                Password = "48334337salamin"
            End If
        End With
    Else
        Password = "48334337salamin"
    End If
        
    key() = StrConv(Password, vbFromUnicode)
    
    For X = 0 To 255
        RB(X) = X
    Next X
    
    X = 0
    Y = 0
    Z = 0
    
    For X = 0 To 255
        Y = (Y + RB(X) + key(X Mod Len(Password))) Mod 256
        temp = RB(X)
        RB(X) = RB(Y)
        RB(Y) = temp
    Next X
    
    X = 0
    Y = 0
    Z = 0
    
    Y = (Y + 1) Mod 256
    Z = (Z + RB(Y)) Mod 256
    temp = RB(Y)
    RB(Y) = RB(Z)
    RB(Z) = temp
    
    EncryptByte = value 'value Xor (RB((RB(Y) + RB(Z)) Mod 256))
    
End Function

Public Function Encryption_RC4_EncryptString(Text As String, Optional key As String) As String
'Encrypts a string with RC4 encryption

    Dim ByteArray() As Byte

    'Convert the data into a byte array
    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the byte array
    Call Encryption_RC4_EncryptByte(ByteArray(), key)

    'Convert the byte array back into a string
    Encryption_RC4_EncryptString = StrConv(ByteArray(), vbUnicode)
End Function

Public Sub Encryption_RC4_SetKey(New_Value As String)
'Sets the encryption key for RC4 encryption

    Dim a As Long
    Dim b As Long
    Dim temp As Byte
    Dim key() As Byte
    Dim KeyLen As Long

    'Do nothing if the key is buffered
    If (m_KeyS = New_Value) Then
        Exit Sub
    End If
    
    'Set the new key
    m_KeyS = New_Value

    'Save the password in a byte array
    key() = StrConv(m_KeyS, vbFromUnicode)
    KeyLen = Len(m_KeyS)

    'Initialize s-boxes
    For a = 0 To 255
        m_sBoxRC4(a) = a
    Next a
    For a = 0 To 255
        b = (b + m_sBoxRC4(a) + key(a Mod KeyLen)) Mod 256
        temp = m_sBoxRC4(a)
        m_sBoxRC4(a) = m_sBoxRC4(b)
        m_sBoxRC4(b) = temp
    Next
End Sub
