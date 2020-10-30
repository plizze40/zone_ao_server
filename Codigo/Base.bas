Attribute VB_Name = "Base"
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Public Con As ADODB.Connection
Public Sub CargarDB()
    On Error GoTo errhandler
    'COMPILAR BETA

    Set Con = New ADODB.Connection
    'Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost; DATABASE=fenixao;UID=fenixao;PWD=lsao666; OPTION=3"
    Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=fenixao;" & "UID=root;PWD=asd123; OPTION=3"
    Con.CursorLocation = adUseClient
    Con.Open

    Exit Sub

errhandler:
    Call LogErrorUrgente("Error en CargarDB: " & Err.Description & " String: " & Con.ConnectionString)
    End

End Sub
Public Function ChangePos(UserName As String) As Boolean
    Call WriteVar(CharPath & Name & ".chr", "INIT", "Position", ULLATHORPE.Map & "-" & ULLATHORPE.X & "-" & ULLATHORPE.Y)
End Function
Public Function ChangeBan(ByVal Name As String, ByVal Baneado As Integer) As Boolean
    Dim Orden As String
    'If GetVar(CharPath & Name & ".chr", "FLAGS", "Ban") <> "0" Then
    '    Call SendData(ToIndex, UserIndex, 0, "||El personaje ya se encuentra baneado." & FONTTYPE_INFO)
    '    Exit Sub
    'End If
    If Baneado = 1 Then
        Call WriteVar(CharPath & Name & ".chr", "FLAGS", "Ban", 1)
    Else
        Call WriteVar(CharPath & Name & ".chr", "FLAGS", "Ban", 0)
    End If

End Function
Public Sub SendCharInfo(ByVal UserName As String, UserIndex As Integer)
    Dim Data As String
    '¿Existe el personaje?

    Dim UserFile As String
    UserFile = CharPath & UCase$(UserName) & ".chr"

    If FileExist(UserFile, vbNormal) = False Then Exit Sub
    Data = "CHRINFO" & UserName
    Data = Data & "," & ListaRazas(val(GetVar(UserFile, "INIT", "Raza"))) & "," & ListaClases(val(GetVar(UserFile, "INIT", "Clase"))) & _
           "," & GeneroLetras(val(GetVar(UserFile, "INIT", "Genero"))) & ","
    Data = Data & val(GetVar(UserFile, "STATS", "ELV")) & "," & val(GetVar(UserFile, "STATS", "GLD")) & "," & val(GetVar(UserFile, "STATS", "BANCO")) & ","

    Data = Data & val(GetVar(UserFile, "Guild", "FundoClan")) & _
           "," & GetVar(UserFile, "Guild", "ClanFundado") & "," _
         & val(GetVar(UserFile, "Guild", "Solicitudes")) & "," _
         & val(GetVar(UserFile, "Guild", "SolicitudesRechazadas")) & "," _
         & val(GetVar(UserFile, "Guild", "VecesFueGuildLeader")) & "," _
         & val(GetVar(UserFile, "Guild", "ClanesParticipo")) & ","

    Data = Data & val(GetVar(UserFile, "FACCIONES", "Bando")) & "," & val(GetVar(UserFile, "FACCIONES", "Matados0")) & "," & val(GetVar(UserFile, "FACCIONES", "Matados1")) & "," & val(GetVar(UserFile, "FACCIONES", "Matados2"))
    Call SendData(ToIndex, UserIndex, 0, Data)

End Sub
Public Sub CerrarDB()
    On Error GoTo ErrHandle

    Con.Close
    Set Con = Nothing

    Exit Sub

ErrHandle:
    Call LogErrorUrgente("Ha surgido un error al cerrar la base de datos MySQL")
    End

End Sub
Public Sub SaveUserSQL(UserIndex As Integer)
    On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String

    mUser = UserList(UserIndex)

    If Len(mUser.Name) = 0 Then Exit Sub

    Set RS = New ADODB.Recordset

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)

    If RS.BOF Or RS.EOF Then
        Con.Execute ("INSERT INTO `charflags` (NOMBRE) VALUES ('" & UCase$(mUser.Name) & "')")
        Set RS = Nothing
        Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(mUser.Name) & "'")
        UserList(UserIndex).IndexPJ = RS!IndexPJ
    End If

    Set RS = Nothing
    Dim Pena As Integer

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    str = "UPDATE `charflags` SET"
    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    str = str & ",Nombre='" & UCase$(mUser.Name) & "'"
    str = str & ",Ban=" & mUser.flags.Ban
    str = str & ",Navegando=" & mUser.flags.Navegando
    str = str & ",Envenenado=" & mUser.flags.Envenenado
    Pena = CalcularTiempoCarcel(UserIndex)
    str = str & ",Pena=" & Pena
    str = str & ",Password='" & mUser.PassWord & "'"
    str = str & ",Canje=" & mUser.flags.Canje    '"
    str = str & ",DenunciasCheat=" & mUser.flags.Denuncias
    str = str & ",DenunciasInsulto=" & mUser.flags.DenunciasInsultos
    str = str & ",EsConseCaos=" & mUser.flags.EsConseCaos
    str = str & ",EsConseReal=" & mUser.flags.EsConseReal
    str = str & ",EsConcilioNegro=" & mUser.flags.EsConcilioNegro
    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charfaccion` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")
    Set RS = Nothing

    str = "UPDATE `charfaccion` SET"

    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    str = str & ",Bando=" & mUser.Faccion.Bando
    str = str & ",BandoOriginal=" & mUser.Faccion.BandoOriginal
    str = str & ",Matados0=" & mUser.Faccion.Matados(0)
    str = str & ",Matados1=" & mUser.Faccion.Matados(1)
    str = str & ",Matados2=" & mUser.Faccion.Matados(2)
    str = str & ",Jerarquia=" & mUser.Faccion.Jerarquia
    str = str & ",Ataco1=" & Buleano(mUser.Faccion.Ataco(1) = 1)
    str = str & ",Ataco2=" & Buleano(mUser.Faccion.Ataco(2) = 1)
    str = str & ",Quests=" & mUser.Faccion.Quests
    str = str & ",Torneos=" & mUser.Faccion.torneos
    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charguild` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")
    Set RS = Nothing

    str = "UPDATE `charguild` SET"

    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    str = str & ",Echadas=" & mUser.GuildInfo.echadas
    str = str & ",SolicitudesRechazadas=" & mUser.GuildInfo.SolicitudesRechazadas
    str = str & ",Guildname='" & mUser.GuildInfo.GuildName & "'"
    str = str & ",ClanesParticipo=" & mUser.GuildInfo.ClanesParticipo
    str = str & ",Guildpts=" & mUser.GuildInfo.GuildPoints
    str = str & ",EsGuildLeader=" & mUser.GuildInfo.EsGuildLeader
    str = str & ",Solicitudes=" & mUser.GuildInfo.Solicitudes
    str = str & ",VecesFueGuildLeader=" & mUser.GuildInfo.VecesFueGuildLeader
    str = str & ",YaVoto=" & mUser.GuildInfo.YaVoto
    str = str & ",FundoClan=" & mUser.GuildInfo.FundoClan
    str = str & ",ClanFundado='" & mUser.GuildInfo.ClanFundado & "'"
    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charatrib` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")
    Set RS = Nothing

    str = "UPDATE `charatrib` SET"
    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    For i = 1 To NUMATRIBUTOS
        str = str & ",AT" & i & "=" & mUser.Stats.UserAtributosBackUP(i)
    Next i
    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charskills` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")
    Set RS = Nothing

    str = "UPDATE `charskills` SET"
    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ

    For i = 1 To NUMSKILLS
        str = str & ",SK" & i & "=" & mUser.Stats.UserSkills(i)
    Next i

    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charinit` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")
    Set RS = Nothing

    str = "UPDATE `charinit` SET"
    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    str = str & ",Email='" & mUser.Email & "'"
    str = str & ",Genero=" & mUser.Genero
    str = str & ",Raza=" & mUser.Raza
    str = str & ",Hogar=" & mUser.Hogar
    str = str & ",Clase=" & mUser.Clase
    str = str & ",Codigo='" & mUser.codigo & "'"
    str = str & ",Descripcion='" & mUser.Desc & "'"
    str = str & ",Head=" & mUser.OrigChar.Head
    str = str & ",LastIP='" & mUser.ip & "'"
    str = str & ",LastMac=" & UserList(UserIndex).Mac & "'"
    str = str & ",Mapa=" & mUser.POS.Map
    str = str & ",X=" & mUser.POS.X
    str = str & ",Y=" & mUser.POS.Y
    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charstats` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")
    Set RS = Nothing

    str = "UPDATE `charstats` SET"
    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    str = str & ",GLD=" & mUser.Stats.GLD
    str = str & ",BANCO=" & mUser.Stats.Banco
    str = str & ",MaxHP=" & mUser.Stats.MaxHP
    str = str & ",MinHP=" & mUser.Stats.MinHP
    str = str & ",MaxMAN=" & mUser.Stats.MaxMAN
    str = str & ",MinMAN=" & mUser.Stats.MinMAN
    str = str & ",MinSTA=" & mUser.Stats.MinSta
    str = str & ",MaxHIT=" & mUser.Stats.MaxHit
    str = str & ",MinHIT=" & mUser.Stats.MinHit
    str = str & ",MinAGU=" & mUser.Stats.MinAGU
    str = str & ",MinHAM=" & mUser.Stats.MinHam
    str = str & ",SkillPtsLibres=" & mUser.Stats.SkillPts
    str = str & ",VecesMurioUsuario=" & mUser.Stats.VecesMurioUsuario
    str = str & ",EXP=" & mUser.Stats.Exp
    str = str & ",ELV=" & mUser.Stats.ELV
    str = str & ",NpcsMuertes=" & mUser.Stats.NPCsMuertos
    For i = 1 To 3
        str = str & ",Recompensa" & i & "=" & mUser.Recompensas(i)
    Next i
    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charbanco` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")

    str = "UPDATE `charbanco` SET"
    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        str = str & ",OBJ" & i & "=" & mUser.BancoInvent.Object(i).OBJIndex
        str = str & ",CANT" & i & "=" & mUser.BancoInvent.Object(i).Amount
    Next i
    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charhechizos` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")
    Set RS = Nothing

    str = "UPDATE `charhechizos` SET"
    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    For i = 1 To MAXUSERHECHIZOS
        str = str & ",H" & i & "=" & mUser.Stats.UserHechizos(i)
    Next i
    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)


    Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & UserList(UserIndex).IndexPJ)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `charinvent` (IndexPJ) VALUES (" & UserList(UserIndex).IndexPJ & ")")
    Set RS = Nothing

    str = "UPDATE `charinvent` SET"
    str = str & " IndexPJ=" & UserList(UserIndex).IndexPJ
    For i = 1 To MAX_INVENTORY_SLOTS
        str = str & ",OBJ" & i & "=" & mUser.Invent.Object(i).OBJIndex
        str = str & ",CANT" & i & "=" & mUser.Invent.Object(i).Amount
    Next i
    str = str & ",CASCOSLOT=" & mUser.Invent.CascoEqpSlot
    str = str & ",ARMORSLOT=" & mUser.Invent.ArmourEqpSlot
    str = str & ",SHIELDSLOT=" & mUser.Invent.EscudoEqpSlot
    str = str & ",WEAPONSLOT=" & mUser.Invent.WeaponEqpSlot
    str = str & ",HERRAMIENTASLOT=" & mUser.Invent.HerramientaEqpslot
    str = str & ",MUNICIONSLOT=" & mUser.Invent.MunicionEqpSlot
    str = str & ",BARCOSLOT=" & mUser.Invent.BarcoSlot

    str = str & " WHERE IndexPJ=" & UserList(UserIndex).IndexPJ
    Call Con.Execute(str)

    Call RevisarTops(UserIndex)

    Exit Sub

ErrHandle:
    Resume Next
End Sub
Function CalcularTiempoCarcel(UserIndex As Integer) As Integer

    If UserList(UserIndex).flags.Encarcelado = 1 Then CalcularTiempoCarcel = 1 + (UserList(UserIndex).Counters.TiempoPena - TiempoTranscurrido(UserList(UserIndex).Counters.Pena)) \ 60

End Function
Function LoadUserSQL(UserIndex As Integer, ByVal Name As String) As Boolean
    On Error GoTo errhandler
    Dim i As Integer

    With UserList(UserIndex)
        Dim RS As New ADODB.Recordset
        Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If

        .IndexPJ = RS!IndexPJ
        Set RS = Nothing

        Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If

        .flags.Ban = RS!Ban
        .flags.Navegando = RS!Navegando
        .flags.Envenenado = RS!Envenenado
        .Counters.TiempoPena = RS!Pena * 60
        .PassWord = RS!PassWord
        .flags.Denuncias = RS!DenunciasCheat
        .flags.DenunciasInsultos = RS!DenunciasInsulto
        .flags.EsConseCaos = RS!EsConseCaos
        .flags.EsConseReal = RS!EsConseReal
        .flags.EsConcilioNegro = RS!EsConcilioNegro
        .flags.Canje = RS!Canje

        Set RS = Nothing


        Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & .IndexPJ)

        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If

        .Faccion.Bando = RS!Bando
        .Faccion.BandoOriginal = RS!BandoOriginal
        .Faccion.Matados(0) = RS!matados0
        .Faccion.Matados(1) = RS!matados1
        .Faccion.Matados(2) = RS!matados2
        .Faccion.Jerarquia = RS!Jerarquia
        .Faccion.Ataco(1) = RS!Ataco1
        .Faccion.Ataco(2) = RS!Ataco2
        .Faccion.Quests = RS!Quests
        .Faccion.torneos = RS!torneos
        Set RS = Nothing

        If Not ModoQuest And UserList(UserIndex).Faccion.Bando <> Neutral And UserList(UserIndex).Faccion.Bando <> UserList(UserIndex).Faccion.BandoOriginal Then UserList(UserIndex).Faccion.Bando = Neutral

        Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If

        .GuildInfo.EsGuildLeader = RS!EsGuildLeader
        .GuildInfo.echadas = RS!echadas
        .GuildInfo.Solicitudes = RS!Solicitudes
        .GuildInfo.SolicitudesRechazadas = RS!SolicitudesRechazadas
        .GuildInfo.VecesFueGuildLeader = RS!VecesFueGuildLeader
        .GuildInfo.YaVoto = RS!YaVoto
        .GuildInfo.FundoClan = RS!FundoClan
        .GuildInfo.GuildName = RS!GuildName
        .GuildInfo.ClanFundado = RS!ClanFundado
        .GuildInfo.ClanesParticipo = RS!ClanesParticipo
        .GuildInfo.GuildPoints = RS!GuildPts
        Set RS = Nothing


        Set RS = Con.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If

        For i = 1 To NUMATRIBUTOS
            .Stats.UserAtributos(i) = RS.Fields("AT" & i)
            .Stats.UserAtributosBackUP(i) = .Stats.UserAtributos(i)
        Next i

        Set RS = Nothing


        Set RS = Con.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = RS.Fields("SK" & i)
        Next i
        Set RS = Nothing


        Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            .BancoInvent.Object(i).OBJIndex = RS.Fields("OBJ" & i)
            .BancoInvent.Object(i).Amount = RS.Fields("CANT" & i)
        Next i
        Set RS = Nothing


        Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
        For i = 1 To MAX_INVENTORY_SLOTS
            .Invent.Object(i).OBJIndex = RS.Fields("OBJ" & i)
            .Invent.Object(i).Amount = RS.Fields("CANT" & i)
        Next i
        .Invent.CascoEqpSlot = RS!CASCOSLOT
        .Invent.ArmourEqpSlot = RS!ARMORSLOT
        .Invent.EscudoEqpSlot = RS!SHIELDSLOT
        .Invent.WeaponEqpSlot = RS!WEAPONSLOT
        .Invent.HerramientaEqpslot = RS!HERRAMIENTASLOT
        .Invent.MunicionEqpSlot = RS!MUNICIONSLOT
        .Invent.BarcoSlot = RS!BarcoSlot
        Set RS = Nothing


        Set RS = Con.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
        For i = 1 To MAXUSERHECHIZOS
            .Stats.UserHechizos(i) = RS.Fields("H" & i)
        Next i
        Set RS = Nothing

        Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
        .Stats.GLD = RS!GLD
        .Stats.Banco = RS!Banco
        .Stats.MaxHP = RS!MaxHP
        .Stats.MinHP = RS!MinHP
        .Stats.MinSta = RS!MinSta
        .Stats.MaxMAN = RS!MaxMAN
        .Stats.MinMAN = RS!MinMAN
        .Stats.MaxHit = RS!MaxHit
        .Stats.MinHit = RS!MinHit
        .Stats.MinAGU = RS!MinAGU
        .Stats.MinHam = RS!MinHam
        .Stats.SkillPts = RS!SkillPtsLibres
        .Stats.VecesMurioUsuario = RS!VecesMurioUsuario
        .Stats.Exp = RS!Exp
        .Stats.ELV = RS!ELV
        .Stats.ELU = ELUs(.Stats.ELV)
        .Stats.NPCsMuertos = RS!NpcsMuertes

        For i = 1 To 3
            .Recompensas(i) = RS.Fields("Recompensa" & i)
        Next

        Set RS = Nothing

        If .Stats.MinAGU < 1 Then .flags.Sed = 1
        If .Stats.MinHam < 1 Then .flags.Hambre = 1
        If .Stats.MinHP < 1 Then .flags.Muerto = 1


        Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & .IndexPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
        .Email = RS!Email
        .Genero = RS!Genero
        .Raza = RS!Raza
        .Hogar = RS!Hogar
        .Clase = RS!Clase
        .codigo = RS!codigo
        .Desc = RS!Descripcion
        .OrigChar.Head = RS!Head
        .POS.Map = RS!mapa
        .POS.X = RS!X
        .POS.Y = RS!Y

        If .flags.Muerto = 0 Then
            .Char = .OrigChar
            Call VerObjetosEquipados(UserIndex)
        Else
            .Char.Body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
        End If

        .Char.Heading = 3

        Set RS = Nothing

        LoadUserSQL = True


        If Len(.Desc) >= 80 Then .Desc = Left$(.Desc, 80)

        If .Counters.TiempoPena > 0 Then
            .flags.Encarcelado = 1
            .Counters.Pena = Timer
        End If

        .Stats.MaxAGU = 100
        .Stats.MaxHam = 100
        Call CalcularSta(UserIndex)

    End With

    Exit Function

errhandler:
    Call LogError("Error en LoadUserSQL. N:" & Name & " - " & Err.Number & "-" & Err.Description)
    Set RS = Nothing

End Function
Function SumarDenuncia(ByVal Name As String, Tipo As Byte) As Integer
    Dim RS As New ADODB.Recordset
    On Error GoTo Error
    Dim str As String, Den As Integer

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

    If RS.BOF Or RS.EOF Then Exit Function

    str = "UPDATE `charflags` SET"
    str = str & " IndexPJ=" & RS!IndexPJ
    str = str & ",Nombre='" & RS!Nombre & "'"
    str = str & ",Ban=" & RS!Ban
    str = str & ",Navegando=" & RS!Navegando
    str = str & ",Envenenado=" & RS!Envenenado
    str = str & ",Pena=" & RS!Pena
    str = str & ",Password='" & RS!PassWord & "'"

    If Tipo = 1 Then
        Den = RS!DenunciasCheat
        SumarDenuncia = Den + 1
        str = str & ",DenunciasCheat=" & SumarDenuncia
        str = str & ",DenunciasInsulto=" & RS!DenunciasInsulto
    Else
        Den = RS!DenunciasInsulto
        SumarDenuncia = Den + 1
        str = str & ",DenunciasCheat=" & RS!DenunciasCheat
        str = str & ",DenunciasInsulto=" & SumarDenuncia
    End If

    str = str & " WHERE IndexPJ=" & RS!IndexPJ
    Call Con.Execute(str)

    Set RS = Nothing
    Exit Function
Error:
    Call LogError("Error en SumarDenuncia: " & Err.Description & " " & Name & " " & Tipo)

End Function
Function ComprobarPassword(ByVal Name As String, PassWord As String, Optional Maestro As Boolean) As Byte
    Dim pass As String

    Dim RS As New ADODB.Recordset

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

    If RS.BOF Or RS.EOF Then Exit Function

    pass = RS!PassWord
    If Len(pass) = 0 Then Exit Function
    Set RS = Nothing

    ComprobarPassword = PassWord = pass

End Function
Public Function BANCheck(ByVal Name As String) As Boolean
'If Inbaneable(Name) Then Exit Function

    BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)    'Or _
                                                                                             (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "AdminBan")) = 1)

End Function
Public Function IndexPJ(ByVal Name As String) As Integer
    Dim RS As New ADODB.Recordset
    Dim Baneado As Byte

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

    If RS.BOF Or RS.EOF Then Exit Function

    IndexPJ = RS!IndexPJ

    Set RS = Nothing

End Function
Function ExistePersonaje(Name As String) As Boolean
    Dim RS As New ADODB.Recordset

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

    If RS.BOF Or RS.EOF Then Exit Function

    Set RS = Nothing

    ExistePersonaje = True

End Function
Function AgregarAClan(ByVal Name As String, ByVal Clan As String) As Boolean
    Dim RS As New ADODB.Recordset
    Dim IndexPJ As Long
    Dim str As String

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

    If RS.BOF Or RS.EOF Then Exit Function

    IndexPJ = RS!IndexPJ

    Set RS = Nothing

    Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & IndexPJ)
    If RS.BOF Or RS.EOF Then Exit Function

    If Len(RS!GuildName) = 0 Then
        str = "UPDATE `charguild` SET"
        str = str & " IndexPJ=" & IndexPJ
        str = str & ",Echadas=" & RS!echadas
        str = str & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas
        str = str & ",Guildname='" & Clan & "'"
        str = str & ",ClanesParticipo=" & RS!ClanesParticipo + 1
        str = str & ",Guildpts=" & RS!GuildPts + 25
        str = str & " WHERE IndexPJ=" & IndexPJ
        Call Con.Execute(str)
        AgregarAClan = True
    End If

    Set RS = Nothing

End Function
Sub RechazarSolicitud(ByVal Name As String)
    Dim RS As New ADODB.Recordset
    Dim IndexPJ As Long
    Dim Orden As String

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

    If RS.BOF Or RS.EOF Then Exit Sub

    IndexPJ = RS!IndexPJ

    Set RS = Nothing

    Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & IndexPJ)
    If RS.BOF Or RS.EOF Then Exit Sub

    Orden = "UPDATE `charguild` SET"
    Orden = Orden & " IndexPJ=" & IndexPJ
    Orden = Orden & ",Echadas=" & RS!echadas
    Orden = Orden & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas + 1
    Orden = Orden & " WHERE IndexPJ=" & IndexPJ
    Call Con.Execute(Orden)

    Set RS = Nothing

End Sub
Sub EcharDeClan(ByVal Name As String)
    Dim RS As New ADODB.Recordset
    Dim IndexPJ As Long
    Dim str As String
    Dim Echa As Integer

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

    If RS.BOF Or RS.EOF Then Exit Sub

    IndexPJ = RS!IndexPJ

    Set RS = Nothing

    Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & IndexPJ)
    If RS.BOF Or RS.EOF Then Exit Sub

    str = "UPDATE `charguild` SET"
    str = str & " IndexPJ=" & IndexPJ
    Echa = RS!echadas
    Echa = Echa + 1
    str = str & ",Echadas=" & Echa
    str = str & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas
    str = str & ",Guildname=''"
    str = str & " WHERE IndexPJ=" & IndexPJ

    Call Con.Execute(str)

    Set RS = Nothing

End Sub

Public Function ComprobarCorreo(ByVal Name As String, Correo As String) As Boolean

    ComprobarCorreo = GetVar(App.Path & "\Charfile\" & Name & ".chr", "CONTACTO", "Email") = Correo

End Function
Public Function ObtenerPassword(ByVal Name As String, Optional PassWord As String) As String

    Dim NuevaPassword As String

    NuevaPassword = Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122))

    Call WriteVar(App.Path & "\Charfile\" & Name & ".chr", "INIT", "Password", MD5String(NuevaPassword))

    ObtenerPassword = NuevaPassword

End Function
