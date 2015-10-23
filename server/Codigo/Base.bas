Attribute VB_Name = "Base"
Option Explicit

'Database connection information (values specified in /ServerData/Server.ini)
Public DB_User As String    'The database username - (default "root")
Public DB_Pass As String    'Password to your database for the corresponding username
Public DB_Name As String    'Name of the table in the database (default "vbgore")
Public DB_Host As String    'IP of the database server - use localhost if hosted locally! Only host remotely for multiple servers!
Public DB_Port As Integer   'Port of the database (default "3306")

Public DB_Conn As New ADODB.Connection
Public DB_RS As New ADODB.Recordset

Public Const OptimizeDatabase As Boolean = False

Public Sub MySQL_Init()

On Error GoTo ErrOut

    Dim ErrorString As String
    Dim i As Long
    
    DB_User = GetVar(ServidorIni, "MYSQL", "User")
    DB_Pass = GetVar(ServidorIni, "MYSQL", "Password")
    DB_Name = GetVar(ServidorIni, "MYSQL", "Database")
    DB_Host = GetVar(ServidorIni, "MYSQL", "Host")
    DB_Port = Val(GetVar(ServidorIni, "MYSQL", "Port"))
    
    'Create the connection
    DB_Conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & DB_Host & _
        ";DATABASE=" & DB_Name & ";PORT=" & DB_Port & ";UID=" & DB_User & ";PWD=" & DB_Pass & ";OPTION=3"
    DB_Conn.CursorLocation = adUseClient
    
    DB_Conn.Open
    
    'Run test queries to make sure the tables are there
    'Call DB_RS_Open ("SELECT * FROM banned_ips WHERE 0=1")
    'DB_RS_Close
    'Call DB_RS_Open ("SELECT * FROM mail WHERE 0=1")
    'DB_RS_Close
    'Call DB_RS_Open ("SELECT * FROM mail_lastid WHERE 0=1")
    'DB_RS_Close
    'Call DB_RS_Open ("SELECT * FROM npcs WHERE 0=1")
    'DB_RS_Close
    'Call DB_RS_Open ("SELECT * FROM objs WHERE 0=1")
    'DB_RS_Close
    'Call DB_RS_Open ("SELECT * FROM quests WHERE 0=1")
    'DB_RS_Close
    
    Call DB_RS_Open("SELECT 1 from people WHERE 0=1")
    DB_RS_Close

    On Error GoTo 0
    
    Exit Sub
    
ErrOut:
    
    'Refresh the errors
    DB_Conn.Errors.Refresh
    
    'Get the error string if there is one
    If DB_Conn.Errors.Count > 0 Then
        ErrorString = DB_Conn.Errors.Item(0).description
    End If
    
    'Check for known errors
    If InStr(1, ErrorString, "Access denied for user ") Then
        'Invalid username or password
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "An incorrect username and/or password was entered into the configuration file." & vbNewLine & _
            "This information can be changed in the Servidor.ini file on the 'User='and 'Password='lines." & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            "Username: " & DB_User & "   Password: " & DB_Pass & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine & _
            "MySQL returned the following error Message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"

    ElseIf InStr(1, ErrorString, "Can't connect to MySQL server on ") Then
        'Unable to connect to MySQL
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "Either an invalid MySQL server IP and/or port was entered, or the server isn't running!" & vbNewLine & _
            "Please confirm you installed MySQL 5.0 and ran the Instance Configuration." & vbNewLine & _
            "To manually start the instance, do the following:" & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            "Right-click 'My Computer'-> 'Manage'-> 'Services and Applications'-> 'Services'" & vbNewLine & _
            "Find your MySQL service in this list (name usually starts with 'MySQL'), right-click it and click 'Start'" & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine & _
            "MySQL returned the following error Message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"
            
    ElseIf InStr(1, ErrorString, "Unknown database ") Then
        'Invalid database name / database does not exist
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "An invalid or unknown database name, '" & DB_Name & "', was entered." & vbNewLine & _
            "This information can be changed in the Servidor.ini file on the 'Database='line." & vbNewLine & vbNewLine & _
            "MySQL returned the following error Message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"
            
    ElseIf InStr(1, ErrorString, "Data source name not found and no default driver specified") Then
        'Invalid database name / database does not exist
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "No valid driver could be found on this computer to connect to MySQL." & vbNewLine & _
            "Please make sure you install ODBC v3.51 (must be v3.51) on this computer!" & vbNewLine & _
            "ODBC can be downloaded from:" & vbNewLine & _
            "http://dev.mysql.com/downloads/connector/odbc/3.51.html" & vbNewLine & vbNewLine & _
            "MySQL returned the following error Message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"
            
    ElseIf InStr(1, ErrorString, "Table '") & InStr(1, ErrorString, "'doesn't exist") Then
        'At least one of the tables are missing
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "One or more of the tables required were not found." & vbNewLine & _
            "Please make sure you import the 'vbgore.sql'file found in the folder '/_Database Dump/'into the database." & vbNewLine & vbNewLine & _
            "MySQL returned the following error Message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"
    
    Else
        'Unknown error
        MsgBox "Unknown error connecting to the MySQL database!" & vbNewLine & _
            "Please confirm that you have completed the following tasks:" & vbNewLine & vbNewLine & _
            " - You have followed ALL of the steps on the MySQL Setup page on the vbGORE site" & vbNewLine & _
            " - MySQL is running and you can connect to it through a GUI such as SQLyog" & vbNewLine & _
            " - You have imported the vbgore.sql file into the database and can see the information through the MySQL GUI" & vbNewLine & _
            " - You have version 5.0 of MySQL and 3.51 of ODBC being used" & vbNewLine & _
            " - You changed the Servidor.ini file to use your MySQL information" & vbNewLine & vbNewLine & _
            "If you are positive you have done all of the above, ask for help on the vbGORE forums." & vbNewLine & vbNewLine & _
            "MySQL returned the following error Message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------", vbOKOnly
    End If
    
    End
    
End Sub

Public Sub MySQL_Optimize()
'Sends a query to the database requesting all tables to be optimized

    'Optimize the database tables
    DB_Conn.Execute "OPTIMIZE TABLE mail, mail_lastid, npcs, objs, quests, users"
End Sub

Public Sub DB_RS_Open(ByVal Query As String)

On Error GoTo Error
1
    DB_RS.Open Query, DB_Conn, adOpenStatic, adLockOptimistic
            
    Exit Sub
    
Error:
    DB_RS_Close
    
    GoTo 1
End Sub

Public Sub DB_RS_Close()

On Error Resume Next
    DB_RS.Close
    
End Sub

Public Sub OnlinePlayers()

    Call DB_RS_Open("SELECT * FROM stats")
    
    DB_RS!Online_Players = Poblacion
    DB_RS.Update
    
    DB_RS_Close

End Sub

Public Function User_Exist(ByVal UserName As String) As Boolean

    Call DB_RS_Open("SELECT name FROM people WHERE `name`='" & UserName & "'")

    User_Exist = Not DB_RS.EOF
    
    DB_RS_Close
    
End Function

Public Function Check_Password(ByVal UserName As String, ByVal Password As String) As Boolean

    Call DB_RS_Open("SELECT * FROM people WHERE `name`='" & UserName & "'")
    
    Dim Pass As String
    Pass = DB_RS!Pass
    
    DB_RS_Close
    
    Check_Password = (Password = Pass)
    
End Function

Public Function Ban_Check(ByVal UserName As String) As Boolean
    
    Call DB_RS_Open("SELECT * FROM people WHERE `name`='" & UserName & "'AND `ban`='1'")
    
    Ban_Check = Not DB_RS.EOF
        
    DB_RS_Close
        
End Function

Public Function SaveItem(ByVal ObjIndex As Integer) As Boolean
    
    'Open the database with an empty record and create the new user
    Call DB_RS_Open("SELECT * FROM items WHERE 0=1")

    DB_RS.AddNew
    
    With ObjData(ObjIndex)
        
        'Put the data in the recordset
        'DB_RS!Name = .Name
        'DB_RS!price = .flags.Password
        'DB_RS!ObjType = Act_Code
        'DB_RS!weapontype = .Email
        'DB_RS!weaponrange = .Raza
        'DB_RS!classreq = .Clase
        'DB_RS!GrhIndex = .Genero
        'DB_RS!usegrh = .Hogar
        'DB_RS!usesfx = .Char.Head
        'DB_RS!projectilerotatespeed = .Stats.Atributos(1)
        'DB_RS!stacking = .Stats.Atributos(2)
        'DB_RS!sprite_body = .Stats.Atributos(3)
        'DB_RS!sprite_weapon = .Stats.Atributos(4)
        'DB_RS!sprite_hair = .Stats.Atributos(5)
        'DB_RS!sprite_head = .Stats.Elv
        'DB_RS!sprite_wings = .Stats.Exp
        'DB_RS!replenish_hp = KSStr
        'DB_RS!replenish_mp = .Skills.NroFree
        'DB_RS!replenish_sp = SpellsStr
        'DB_RS!replenish_hp_percent = .Pos.Map
        'DB_RS!replenish_mp_percent = .Pos.X
        'DB_RS!replenish_sp_percent = .Pos.Y
        'DB_RS!stat_str = .Stats.MinHP
        'DB_RS!stat_agi = .Stats.MaxHP
        'DB_RS!stat_mag = .Stats.MinMan
        'DB_RS!stat_def = .Stats.MaxMan
        'DB_RS!stat_speed = .Stats.MinSta
        'DB_RS!stat_hit_min = .Stats.MaxSta
        'DB_RS!stat_hit_max = .Stats.MinHit
        'DB_RS!stat_hp = .Stats.MaxHit
        'DB_RS!stat_mp = .Stats.MinSed
        'DB_RS!stat_sp = .Stats.MinHam
        'DB_RS!req_str = .InvStr
        'DB_RS!req_agi = .BeltStr
        'DB_RS!req_mag = .BankStr
        'DB_RS!req_lvl = .req_lvl
        
    End With
    
    DB_RS_Close
        
End Function

Public Function SaveSpell(ByVal SpellIndex As Integer) As Boolean
    
    'Open the database with an empty record and create the new user
    Call DB_RS_Open("SELECT * FROM spells WHERE 0=1")

    DB_RS.AddNew
    
    With Hechizos(SpellIndex)
        
        'Put the data in the recordset
        'DB_RS!Name = .Name
        'DB_RS!Desc = .flags.Password
        'DB_RS!magicWords = Act_Code
        'DB_RS!HechizeroMsg = .Email
        'DB_RS!PropioMsg = .Raza
        'DB_RS!TargetMsg = .Clase
        'DB_RS!Type = .Genero
        'DB_RS!snd = .Hogar
        'DB_RS!FXgrh = .Char.Head
        'DB_RS!FXLoops = .Stats.Atributos(1)
        'DB_RS!MinSkill = .Stats.Atributos(2)
        'DB_RS!manaRequired = .Stats.Atributos(3)
        'DB_RS!staRequired = .Stats.Atributos(4)
        'DB_RS!TargetType = .Stats.Atributos(5)
        'DB_RS!SubeHP = .Stats.Elv
        'DB_RS!MinHP = .Stats.Exp
        'DB_RS!MaxHP = KSStr
        'DB_RS!SubeMana = .Skills.NroFree
        'DB_RS!minMana = SpellsStr
        'DB_RS!maxMana = .Pos.Map
        'DB_RS!SubeSta = .Pos.X
        'DB_RS!MinSta = .Pos.Y
        'DB_RS!MaxSta = .Stats.MinHP
        'DB_RS!SubeHam = .Stats.MaxHP
        'DB_RS!MinHam = .Stats.MinMan
        'DB_RS!MaxHam = .Stats.MaxMan
        'DB_RS!SubeSed = .Stats.MinSta
        'DB_RS!MinSed = .Stats.MaxSta
        'DB_RS!MaxSed = .Stats.MinHit
        'DB_RS!subeAg = .Stats.MaxHit
        'DB_RS!minAg = .Stats.MinSed
        'DB_RS!maxAg = .Stats.MinHam
        'DB_RS!subeFu = .InvStr
        'DB_RS!minFu = .BeltStr
        'DB_RS!maxFu = .BankStr
        'DB_RS!subeCa = .InvStr
        'DB_RS!minCa = .BeltStr
        'DB_RS!maxCa = .BankStr
        'DB_RS!invi = .InvStr
        'DB_RS!Paraliza = .BeltStr
        'DB_RS!inmo = .BankStr
        'DB_RS!remueveInmo = .InvStr
        'DB_RS!remueveEstupidez = .BeltStr
        'DB_RS!remueveInviParcial = .BankStr
        'DB_RS!CuraVeneno = .InvStr
        'DB_RS!Envenena = .BeltStr
        'DB_RS!revive = .BankStr
        'DB_RS!enceguece = .BankStr
        'DB_RS!Estupidez = .InvStr
        'DB_RS!Invoca = .BeltStr
        'DB_RS!NumNpc = .BankStr
        'DB_RS!cantidadNpc = .InvStr
        'DB_RS!Mimetiza = .BeltStr
        'DB_RS!materializa = .BankStr
        'DB_RS!ItemIndex = .InvStr
        'DB_RS!StaffAffected = .BeltStr
        'DB_RS!NeedStaff = .BankStr
        'DB_RS!resistencia = .InvStr
    End With
    
    DB_RS_Close
        
End Function
