Attribute VB_Name = "modFonts"
Private Type tFont
    Red As Byte
    Green As Byte
    Blue As Byte
    Bold As Boolean
    Italic As Boolean
End Type

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_TALKGM
    FONTTYPE_YELL
    FONTTYPE_YELLGM
    FONTTYPE_PUBLICMESSAGE
    FONTTYPE_COMPAMESSAGE
    FONTTYPE_PRIVATEMESSAGE
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_NUMERO
    FONTTYPE_HABILIDAD
    FONTTYPE_HECHIZO
End Enum

Public FontTypes(25) As tFont

Public Sub InitFonts()
'Initializes the fonts array

    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .Red = 230
        .Green = 230
        .Blue = 230
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TALKGM)
        .Red = 180
        .Green = 255
        .Blue = 180
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_YELL)
        .Red = 255
        .Green = 200
        .Blue = 100
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_YELLGM)
        .Red = 135
        .Green = 255
        .Blue = 135
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PUBLICMESSAGE)
        .Red = 150
        .Green = 150
        .Blue = 150
        .Italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_COMPAMESSAGE)
        .Red = 100
        .Green = 150
        .Blue = 150
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PRIVATEMESSAGE)
        .Red = 100
        .Green = 150
        .Blue = 150
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .Red = 200
        .Bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .Red = 200
        .Green = 200
        .Blue = 0
        .Bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .Red = 225
        .Green = 220
        .Blue = 185
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .Red = 225
        .Green = 220
        .Blue = 185
        .Bold = True
    End With
    
    'With FontTypes(FontTypeNames.FONTTYPE_INFO)
    '    .Red = 200
    '    .Green = 160
    '    .Blue = 70
    'End With
    
    'With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
    '    .Red = 255
    '    .Green = 250
    '    .Blue = 150
    '    .Bold = True
    'End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .Red = 130
        .Green = 130
        .Blue = 130
        .Bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .Red = 255
        .Green = 180
        .Blue = 250
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).Green = 215
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .Red = 255
        .Green = 255
        .Blue = 255
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).Green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .Red = 228
        .Green = 199
        .Blue = 27
        .Italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .Green = 255
        .Bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .Red = 255
        .Green = 255
        .Blue = 255
        .Italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .Red = 128
        .Green = 255
        .Blue = 128
    End With
        
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .Red = 30
        .Green = 150
        .Blue = 30
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .Red = 250
        .Green = 250
        .Blue = 150
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NUMERO)
        .Red = 255
        .Green = 180
        .Blue = 0
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_HABILIDAD)
        .Red = 200
        .Green = 128
        .Blue = 128
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_HECHIZO)
        .Red = 140
        .Green = 60
        .Blue = 160
    End With
    
End Sub
