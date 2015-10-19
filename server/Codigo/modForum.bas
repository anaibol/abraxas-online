Attribute VB_Name = "modForum"
Option Explicit

Public Const Max_MENSAJES_FORO As Byte = 30
Public Const Max_ANUNCIOS_FORO As Byte = 5

Public Type tPost
    sTitulo As String
    sPost As String
    Autor As String
End Type

Public Type tForo
    vsPost(1 To Max_MENSAJES_FORO) As tPost
    vsAnuncio(1 To Max_ANUNCIOS_FORO) As tPost
    CantPosts As Byte
    CantAnuncios As Byte
    ID As String
End Type

Private NumForos As Integer
Private Foros() As tForo


Public Sub AddForum(ByVal sForoID As String)
'Adds a forum to the list and fills it.
    Dim ForumPath As String
    Dim PostPath As String
    Dim PostIndex As Integer
    Dim FileIndex As Integer
    
    NumForos = NumForos + 1
    ReDim Preserve Foros(1 To NumForos) As tForo
    
    ForumPath = App.Path & "/foros/" & sForoID & ".for"
    
    With Foros(NumForos)
    
        .ID = sForoID
        
        If FileExist(ForumPath, vbNormal) Then
            .CantPosts = val(GetVar(ForumPath, "INFO", "CantMSG"))
            .CantAnuncios = val(GetVar(ForumPath, "INFO", "CantAnuncios"))
            
            'Cargo posts
            For PostIndex = 1 To .CantPosts
                FileIndex = FreeFile
                PostPath = App.Path & "/foros/" & sForoID & PostIndex & ".for"

                Open PostPath For Input Shared As #FileIndex
                
                'Titulo
                Input #FileIndex, .vsPost(PostIndex).sTitulo
                'Autor
                Input #FileIndex, .vsPost(PostIndex).Autor
                'Mensaje
                Input #FileIndex, .vsPost(PostIndex).sPost
                
                Close #FileIndex
            Next PostIndex
            
            'Cargo anuncios
            For PostIndex = 1 To .CantAnuncios
                FileIndex = FreeFile
                PostPath = App.Path & "/foros/" & sForoID & PostIndex & "a.for"

                Open PostPath For Input Shared As #FileIndex
                
                'Titulo
                Input #FileIndex, .vsAnuncio(PostIndex).sTitulo
                'Autor
                Input #FileIndex, .vsAnuncio(PostIndex).Autor
                'Mensaje
                Input #FileIndex, .vsAnuncio(PostIndex).sPost
                
                Close #FileIndex
            Next PostIndex
        End If
        
    End With
    
End Sub

Public Function GetForumIndex(ByRef sForoID As String) As Integer

'Author: ZaMa
'Last Modification: 22\02\2010
'Returns the forum Index.

    
    Dim ForumIndex As Integer
    
    For ForumIndex = 1 To NumForos
        If Foros(ForumIndex).ID = sForoID Then
            GetForumIndex = ForumIndex
            Exit Function
        End If
    Next ForumIndex
    
End Function

Public Sub AddPost(ByVal ForumIndex As Integer, ByRef Post As String, ByRef Autor As String, _
                   ByRef Titulo As String, ByVal bAnuncio As Boolean)

'Author: ZaMa
'Last Modification: 22\02\2010
'Saves a new post into the forum.


    With Foros(ForumIndex)
        
        If bAnuncio Then
            If .CantAnuncios < Max_ANUNCIOS_FORO Then _
                .CantAnuncios = .CantAnuncios + 1
            
            Call MoveArray(ForumIndex, bAnuncio)
            
            'Agrego el anuncio
            With .vsAnuncio(1)
                .sTitulo = Titulo
                .Autor = Autor
                .sPost = Post
            End With
            
        Else
            If .CantPosts < Max_MENSAJES_FORO Then _
                .CantPosts = .CantPosts + 1
                
            Call MoveArray(ForumIndex, bAnuncio)
            
            'Agrego el post
            With .vsPost(1)
                .sTitulo = Titulo
                .Autor = Autor
                .sPost = Post
            End With
        
        End If
    End With
End Sub

Public Sub SaveForums()

'Author: ZaMa
'Last Modification: 22\02\2010
'Saves all forums into disk.

    Dim ForumIndex As Integer

    For ForumIndex = 1 To NumForos
        Call SaveForum(ForumIndex)
    Next ForumIndex
End Sub


Private Sub SaveForum(ByVal ForumIndex As Integer)

'Author: ZaMa
'Last Modification: 22\02\2010
'Saves a forum into disk.


    Dim PostIndex As Integer
    Dim FileIndex As Integer
    Dim PostPath As String
    
    Call CleanForum(ForumIndex)
    
    With Foros(ForumIndex)
        
        'Guardo info del foro
        Call WriteVar(App.Path & "/Foros/" & .ID & ".for", "INFO", "CantMSG", .CantPosts)
        Call WriteVar(App.Path & "/Foros/" & .ID & ".for", "INFO", "CantAnuncios", .CantAnuncios)
        
        'Guardo posts
        For PostIndex = 1 To .CantPosts
            
            PostPath = App.Path & "/Foros/" & .ID & PostIndex & ".for"
            FileIndex = FreeFile()
            Open PostPath For Output As FileIndex
            
            With .vsPost(PostIndex)
                Print #FileIndex, .sTitulo
                Print #FileIndex, .Autor
                Print #FileIndex, .sPost
            End With
            
            Close #FileIndex
            
        Next PostIndex
        
        'Guardo Anuncios
        For PostIndex = 1 To .CantAnuncios
            
            PostPath = App.Path & "/Foros/" & .ID & PostIndex & "a.for"
            FileIndex = FreeFile()
            Open PostPath For Output As FileIndex
            
            With .vsAnuncio(PostIndex)
                Print #FileIndex, .sTitulo
                Print #FileIndex, .Autor
                Print #FileIndex, .sPost
            End With
            
            Close #FileIndex

        Next PostIndex
        
    End With
    
End Sub

Public Sub CleanForum(ByVal ForumIndex As Integer)
'Cleans a forum from disk.

    Dim PostIndex As Integer
    Dim NumPost As Integer
    Dim ForumPath As String

    With Foros(ForumIndex)
    
        'Elimino todo
        ForumPath = App.Path & "/Foros/" & .ID & ".for"
        If FileExist(ForumPath, vbNormal) Then
    
            NumPost = val(GetVar(ForumPath, "INFO", "CantMSG"))
            
            'Elimino los post viejos
            For PostIndex = 1 To NumPost
                Kill App.Path & "/Foros/" & .ID & PostIndex & ".for"
            Next PostIndex
            
            
            NumPost = val(GetVar(ForumPath, "INFO", "CantAnuncios"))
            
            'Elimino los post viejos
            For PostIndex = 1 To NumPost
                Kill App.Path & "/Foros/" & .ID & PostIndex & "a.for"
            Next PostIndex
            
            
            'Elimino el foro
            Kill App.Path & "/Foros/" & .ID & ".for"
    
        End If
    End With

End Sub

Public Function SendPosts(ByVal UserIndex As Integer, ByRef ForoID As String) As Boolean

'Author: ZaMa
'Last Modification: 22\02\2010
'Sends all the posts of a required forum

    
    Dim ForumIndex As Integer
    Dim PostIndex As Integer
    Dim bEsGm As Boolean
    
    ForumIndex = GetForumIndex(ForoID)

    If ForumIndex > 0 Then

        With Foros(ForumIndex)
            
            'Send General posts
            For PostIndex = 1 To .CantPosts
                With .vsPost(PostIndex)
                    Call WriteAddForumMsg(UserIndex, eForumMsgType.ieGeneral, .sTitulo, .Autor, .sPost)
                End With
            Next PostIndex
            
            'Send Sticky posts
            For PostIndex = 1 To .CantAnuncios
                With .vsAnuncio(PostIndex)
                    Call WriteAddForumMsg(UserIndex, eForumMsgType.ieGENERAL_STICKY, .sTitulo, .Autor, .sPost)
                End With
            Next PostIndex
            
        End With
        
        bEsGm = EsGM(UserIndex)
        
        'Caos?
        If esCaos(UserIndex) Or bEsGm Then
            
            ForumIndex = GetForumIndex(FORO_CAOS_ID)
            
            With Foros(ForumIndex)
                
                'Send General Caos posts
                For PostIndex = 1 To .CantPosts
                
                    With .vsPost(PostIndex)
                        Call WriteAddForumMsg(UserIndex, eForumMsgType.ieCAOS, .sTitulo, .Autor, .sPost)
                    End With
                    
                Next PostIndex
                
                'Send Sticky posts
                For PostIndex = 1 To .CantAnuncios
                    With .vsAnuncio(PostIndex)
                        Call WriteAddForumMsg(UserIndex, eForumMsgType.ieCAOS_STICKY, .sTitulo, .Autor, .sPost)
                    End With
                Next PostIndex
                
            End With
        End If
            
        'Caos?
        If esArmada(UserIndex) Or bEsGm Then
            
            ForumIndex = GetForumIndex(FORO_REAL_ID)
            
            With Foros(ForumIndex)
                
                'Send General Real posts
                For PostIndex = 1 To .CantPosts
                
                    With .vsPost(PostIndex)
                        Call WriteAddForumMsg(UserIndex, eForumMsgType.ieREAL, .sTitulo, .Autor, .sPost)
                    End With
                    
                Next PostIndex
                
                'Send Sticky posts
                For PostIndex = 1 To .CantAnuncios
                    With .vsAnuncio(PostIndex)
                        Call WriteAddForumMsg(UserIndex, eForumMsgType.ieREAL_STICKY, .sTitulo, .Autor, .sPost)
                    End With
                Next PostIndex
                
            End With
        End If
        
        SendPosts = True
    End If
    
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean

'Author: ZaMa
'Last Modification: 22\02\2010
'Returns true if the post is sticky.

    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte

'Author: ZaMa
'Last Modification: 01\03\2010
'Returns the forum alignment.

    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function

Public Sub ResetForums()

'Author: ZaMa
'Last Modification: 22\02\2010
'Resets forum info

    ReDim Foros(1 To 1) As tForo
    NumForos = 0
End Sub

Private Sub MoveArray(ByVal ForumIndex As Integer, ByVal Sticky As Boolean)
Dim i As Long

With Foros(ForumIndex)
    If Sticky Then
        For i = .CantAnuncios To 2 Step -1
            .vsAnuncio(i).sTitulo = .vsAnuncio(i - 1).sTitulo
            .vsAnuncio(i).sPost = .vsAnuncio(i - 1).sPost
            .vsAnuncio(i).Autor = .vsAnuncio(i - 1).Autor
        Next i
    Else
        For i = .CantPosts To 2 Step -1
            .vsPost(i).sTitulo = .vsPost(i - 1).sTitulo
            .vsPost(i).sPost = .vsPost(i - 1).sPost
            .vsPost(i).Autor = .vsPost(i - 1).Autor
        Next i
    End If
End With
End Sub
