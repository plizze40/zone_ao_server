Attribute VB_Name = "mdlCPUser"
Sub LoadUser(UserIndex As Integer, UserFile As String)
    On Error GoTo errhandler
    Dim LoopC As Integer
    Dim ln As String
    Dim ln2 As String
    'CARGAMOS USER'
    UserList(UserIndex).PassWord = GetVar(UserFile, "INIT", "Password")

    UserList(UserIndex).Email = GetVar(UserFile, "CONTACTO", "Email")
    UserList(UserIndex).Genero = GetVar(UserFile, "INIT", "Genero")
    UserList(UserIndex).Raza = GetVar(UserFile, "INIT", "Raza")
    UserList(UserIndex).Hogar = GetVar(UserFile, "INIT", "Hogar")
    UserList(UserIndex).Clase = GetVar(UserFile, "INIT", "Clase")
    UserList(UserIndex).codigo = GetVar(UserFile, "INIT", "codigo")

    UserList(UserIndex).Desc = GetVar(UserFile, "INIT", "Desc")

    UserList(UserIndex).OrigChar.Head = val(GetVar(UserFile, "INIT", "Head"))

    UserList(UserIndex).POS.Map = val(ReadField(1, GetVar(UserFile, "INIT", "Position"), 45))
    UserList(UserIndex).POS.X = val(ReadField(2, GetVar(UserFile, "INIT", "Position"), 45))
    UserList(UserIndex).POS.Y = val(ReadField(3, GetVar(UserFile, "INIT", "Position"), 45))

    UserList(UserIndex).Char.Heading = 3

    'CARGAMOS STATS'
    UserList(UserIndex).Stats.GLD = val(GetVar(UserFile, "STATS", "GLD"))
    UserList(UserIndex).Stats.Banco = val(GetVar(UserFile, "STATS", "Banco"))

    UserList(UserIndex).Stats.MaxHP = val(GetVar(UserFile, "STATS", "MaxHP"))
    UserList(UserIndex).Stats.MinHP = val(GetVar(UserFile, "STATS", "MinHP"))

    UserList(UserIndex).Stats.MinSta = val(GetVar(UserFile, "STATS", "MinSta"))
    UserList(UserIndex).Stats.MaxSta = val(GetVar(UserFile, "STATS", "MaxSta"))

    UserList(UserIndex).Stats.MaxMAN = val(GetVar(UserFile, "STATS", "MaxMAN"))

    UserList(UserIndex).Stats.MinMAN = val(GetVar(UserFile, "STATS", "MinMAN"))

    UserList(UserIndex).Stats.MaxHit = val(GetVar(UserFile, "STATS", "MaxHit"))
    UserList(UserIndex).Stats.MinHit = val(GetVar(UserFile, "STATS", "MinHit"))

    UserList(UserIndex).Stats.MinAGU = val(GetVar(UserFile, "STATS", "MinAGU"))
    UserList(UserIndex).Stats.MinHam = val(GetVar(UserFile, "STATS", "MinHam"))

    UserList(UserIndex).Stats.SkillPts = val(GetVar(UserFile, "STATS", "SkillPtsLibres"))

    UserList(UserIndex).Stats.Exp = val(GetVar(UserFile, "STATS", "EXP"))
    UserList(UserIndex).Stats.ELV = val(GetVar(UserFile, "STATS", "ELV"))
    UserList(UserIndex).Stats.ELU = ELUs(val(GetVar(UserFile, "STATS", "ELV")))

    UserList(UserIndex).Stats.VecesMurioUsuario = val(GetVar(UserFile, "MUERTES", "VecesMurioUsuario"))
    UserList(UserIndex).Stats.NPCsMuertos = val(GetVar(UserFile, "MUERTES", "NpcsMuertes"))


    For LoopC = 1 To 3
        UserList(UserIndex).Recompensas(LoopC) = val(GetVar(UserFile, "RECOMPENSAS", "Recompensa" & LoopC))
    Next

    With UserList(UserIndex)
        If .Stats.MinAGU < 1 Then .flags.Sed = 1
        If .Stats.MinHam < 1 Then .flags.Hambre = 1
        If .Stats.MinHP < 1 Then .flags.Muerto = 1
    End With

    'CARGAMOS .FLAGS'
    UserList(UserIndex).flags.Ban = val(GetVar(UserFile, "FLAGS", "Ban"))
    UserList(UserIndex).flags.Navegando = val(GetVar(UserFile, "FLAGS", "Navegando"))
    UserList(UserIndex).flags.Envenenado = val(GetVar(UserFile, "FLAGS", "Envenenado"))
    UserList(UserIndex).flags.Denuncias = val(GetVar(UserFile, "FLAGS", "DenunciasCheat"))
    UserList(UserIndex).flags.DenunciasInsultos = val(GetVar(UserFile, "FLAGS", "DenunciasInsultos"))
    UserList(UserIndex).flags.EsConseReal = val(GetVar(UserFile, "FLAGS", "EsConseReal"))
    UserList(UserIndex).flags.EsConseCaos = val(GetVar(UserFile, "FLAGS", "EsConseCaos"))
    UserList(UserIndex).flags.EsConcilioNegro = val(GetVar(UserFile, "FLAGS", "EsConcilioNegro"))
    UserList(UserIndex).flags.Canje = val(GetVar(UserFile, "FLAGS", "Canje"))

    'CARGAMOS COUNTERS
    UserList(UserIndex).Counters.TiempoPena = val(GetVar(UserFile, "COUNTERS", "Pena"))

    'CARGAMOS FACCION
    UserList(UserIndex).Faccion.Bando = val(GetVar(UserFile, "FACCIONES", "Bando"))
    UserList(UserIndex).Faccion.BandoOriginal = val(GetVar(UserFile, "FACCIONES", "BandoOriginal"))
    UserList(UserIndex).Faccion.Matados(0) = val(GetVar(UserFile, "FACCIONES", "Matados0"))
    UserList(UserIndex).Faccion.Matados(1) = val(GetVar(UserFile, "FACCIONES", "Matados1"))
    UserList(UserIndex).Faccion.Matados(2) = val(GetVar(UserFile, "FACCIONES", "Matados2"))
    UserList(UserIndex).Faccion.Jerarquia = val(GetVar(UserFile, "FACCIONES", "Jerarquia"))
    UserList(UserIndex).Faccion.Ataco(1) = val(GetVar(UserFile, "FACCIONES", "Ataco1"))
    UserList(UserIndex).Faccion.Ataco(2) = val(GetVar(UserFile, "FACCIONES", "Ataco2"))
    UserList(UserIndex).Faccion.Quests = val(GetVar(UserFile, "FACCIONES", "Quests"))
    UserList(UserIndex).Faccion.torneos = val(GetVar(UserFile, "FACCIONES", "Torneos"))


    'CARGAMOS GUILD
    UserList(UserIndex).GuildInfo.EsGuildLeader = val(GetVar(UserFile, "Guild", "EsGuildLeader"))
    UserList(UserIndex).GuildInfo.echadas = val(GetVar(UserFile, "Guild", "Echadas"))
    UserList(UserIndex).GuildInfo.Solicitudes = val(GetVar(UserFile, "Guild", "Solicitudes"))
    UserList(UserIndex).GuildInfo.SolicitudesRechazadas = val(GetVar(UserFile, "Guild", "SolicitudesRechazadas"))
    UserList(UserIndex).GuildInfo.VecesFueGuildLeader = val(GetVar(UserFile, "Guild", "VecesFueGuildLeader"))
    UserList(UserIndex).GuildInfo.YaVoto = val(GetVar(UserFile, "Guild", "YaVoto"))
    UserList(UserIndex).GuildInfo.FundoClan = val(GetVar(UserFile, "Guild", "FundoClan"))
    UserList(UserIndex).GuildInfo.ClanesParticipo = val(GetVar(UserFile, "Guild", "ClanesParticipo"))
    UserList(UserIndex).GuildInfo.GuildPoints = val(GetVar(UserFile, "Guild", "GuildPts"))
    UserList(UserIndex).GuildInfo.GuildName = GetVar(UserFile, "Guild", "GuildName")
    UserList(UserIndex).GuildInfo.ClanFundado = GetVar(UserFile, "Guild", "ClanFundado")

    For LoopC = 1 To NUMATRIBUTOS
        UserList(UserIndex).Stats.UserAtributos(LoopC) = GetVar(UserFile, "ATRIBUTOS", "AT" & LoopC)
        UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
    Next

    For LoopC = 1 To NUMSKILLS
        UserList(UserIndex).Stats.UserSkills(LoopC) = val(GetVar(UserFile, "SKILLS", "SK" & LoopC))
    Next

    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = val(GetVar(UserFile, "Hechizos", "H" & LoopC))
    Next

    '[KEVIN]--------------------------------------------------------------------
    '***********************************************************************************
    UserList(UserIndex).BancoInvent.NroItems = val(GetVar(UserFile, "BancoInventory", "CantidadItems"))
    'Lista de objetos del banco
    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        ln2 = GetVar(UserFile, "BancoInventory", "Obj" & loopd)
        UserList(UserIndex).BancoInvent.Object(loopd).OBJIndex = val(ReadField(1, ln2, 45))
        UserList(UserIndex).BancoInvent.Object(loopd).Amount = val(ReadField(2, ln2, 45))
    Next loopd
    '------------------------------------------------------------------------------------
    '[/KEVIN]*****************************************************************************


    'Lista de objetos
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(UserFile, "Inventory", "Obj" & LoopC)
        UserList(UserIndex).Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
        UserList(UserIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
        UserList(UserIndex).Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
    Next LoopC


    UserList(UserIndex).Invent.WeaponEqpSlot = val(GetVar(UserFile, "Inventory", "WEAPONSLOT"))
    UserList(UserIndex).Invent.CascoEqpSlot = val(GetVar(UserFile, "Inventory", "CASCOSLOT"))
    UserList(UserIndex).Invent.ArmourEqpSlot = val(GetVar(UserFile, "Inventory", "ARMORSLOT"))
    UserList(UserIndex).Invent.EscudoEqpSlot = val(GetVar(UserFile, "Inventory", "SHIELDSLOT"))
    UserList(UserIndex).Invent.HerramientaEqpslot = val(GetVar(UserFile, "Inventory", "HERRAMIENTASLOT"))
    UserList(UserIndex).Invent.MunicionEqpSlot = val(GetVar(UserFile, "Inventory", "MUNICIONSLOT"))
    UserList(UserIndex).Invent.BarcoSlot = val(GetVar(UserFile, "Inventory", "BarcoSlot"))


    With UserList(UserIndex)
        If Len(.Desc) >= 80 Then .Desc = Left$(.Desc, 80)

        If .Counters.TiempoPena > 0 Then
            .flags.Encarcelado = 1
            .Counters.Pena = Timer
        End If

        .Stats.MaxAGU = 100
        .Stats.MaxHam = 100
        Call CalcularSta(UserIndex)
    End With

    With UserList(UserIndex)
        If .flags.Muerto = 0 Then
            .Char = .OrigChar
            UserList(UserIndex).Char.Heading = 3
            Call VerObjetosEquipados(UserIndex)
        Else
            .Char.Body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
        End If
    End With

errhandler:
    Call LogError("Error en LoadUser. N:" & Name & " - " & Err.Number & "-" & Err.Description)
End Sub

Sub SaveUser(UserIndex As Integer, UserFile As String)
    On Error GoTo errhandler
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    Dim Pena As Integer

    Dim OldUserHead As Long

    If FileExist(UserFile, vbNormal) Then
        If UserList(UserIndex).flags.Muerto = 1 Then
            OldUserHead = UserList(UserIndex).Char.Head
            UserList(UserIndex).Char.Head = val(GetVar(UserFile, "INIT", "Head"))
        End If
        'Kill UserFile
    End If

    Dim LoopC As Integer

    Call WriteVar(UserFile, "FLAGS", "Ban", val(UserList(UserIndex).flags.Ban))
    Call WriteVar(UserFile, "FLAGS", "Muerto", val(UserList(UserIndex).flags.Muerto))

    Call WriteVar(UserFile, "FLAGS", "Navegando", val(UserList(UserIndex).flags.Navegando))
    Call WriteVar(UserFile, "FLAGS", "Envenenado", val(UserList(UserIndex).flags.Envenenado))

    Call WriteVar(UserFile, "FLAGS", "DenunciasCheat", val(UserList(UserIndex).flags.Denuncias))
    Call WriteVar(UserFile, "FLAGS", "DenunciasInsultos", val(UserList(UserIndex).flags.DenunciasInsultos))

    Call WriteVar(UserFile, "FLAGS", "EsConseReal", val(UserList(UserIndex).flags.EsConseReal))
    Call WriteVar(UserFile, "FLAGS", "EsConseCaos", val(UserList(UserIndex).flags.EsConseCaos))
    Call WriteVar(UserFile, "FLAGS", "EsConcilioNegro", val(UserList(UserIndex).flags.EsConcilioNegro))
    Call WriteVar(UserFile, "FLAGS", "Canje", val(UserList(UserIndex).flags.Canje))

    Pena = CalcularTiempoCarcel(UserIndex)
    'str = str & ",Pena=" & Pena
    Call WriteVar(UserFile, "COUNTERS", "Pena", val(Pena))
    Call WriteVar(UserFile, "INIT", "Password", UserList(UserIndex).PassWord)


    '****************************************************************************************************************
    '******************************** FACCION ***********************************************************************
    '****************************************************************************************************************

    Call WriteVar(UserFile, "FACCIONES", "Bando", val(UserList(UserIndex).Faccion.Bando))
    Call WriteVar(UserFile, "FACCIONES", "BandoOriginal", val(UserList(UserIndex).Faccion.BandoOriginal))
    Call WriteVar(UserFile, "FACCIONES", "Matados0", val(UserList(UserIndex).Faccion.Matados(0)))
    Call WriteVar(UserFile, "FACCIONES", "Matados1", val(UserList(UserIndex).Faccion.Matados(1)))
    Call WriteVar(UserFile, "FACCIONES", "Matados2", val(UserList(UserIndex).Faccion.Matados(2)))

    Call WriteVar(UserFile, "FACCIONES", "Jerarquia", val(UserList(UserIndex).Faccion.Jerarquia))
    Call WriteVar(UserFile, "FACCIONES", "Ataco1", Buleano(UserList(UserIndex).Faccion.Ataco(1) = 1))
    Call WriteVar(UserFile, "FACCIONES", "Ataco2", Buleano(UserList(UserIndex).Faccion.Ataco(2) = 1))

    Call WriteVar(UserFile, "FACCIONES", "Quests", val(UserList(UserIndex).Faccion.Quests))
    Call WriteVar(UserFile, "FACCIONES", "Torneos", val(UserList(UserIndex).Faccion.torneos))

    '****************************************************************************************************************
    '******************************** GUILDS ************************************************************************
    '****************************************************************************************************************

    Call WriteVar(UserFile, "GUILD", "EsGuildLeader", val(UserList(UserIndex).GuildInfo.EsGuildLeader))
    Call WriteVar(UserFile, "GUILD", "Echadas", val(UserList(UserIndex).GuildInfo.echadas))
    Call WriteVar(UserFile, "GUILD", "Solicitudes", val(UserList(UserIndex).GuildInfo.Solicitudes))
    Call WriteVar(UserFile, "GUILD", "SolicitudesRechazadas", val(UserList(UserIndex).GuildInfo.SolicitudesRechazadas))
    Call WriteVar(UserFile, "GUILD", "VecesFueGuildLeader", val(UserList(UserIndex).GuildInfo.VecesFueGuildLeader))
    Call WriteVar(UserFile, "GUILD", "YaVoto", val(UserList(UserIndex).GuildInfo.YaVoto))
    Call WriteVar(UserFile, "GUILD", "FundoClan", val(UserList(UserIndex).GuildInfo.FundoClan))

    Call WriteVar(UserFile, "GUILD", "ClanesParticipo", val(UserList(UserIndex).GuildInfo.ClanesParticipo))
    Call WriteVar(UserFile, "GUILD", "GuildPts", val(UserList(UserIndex).GuildInfo.GuildPoints))
    Call WriteVar(UserFile, "GUILD", "GuildName", UserList(UserIndex).GuildInfo.GuildName)
    Call WriteVar(UserFile, "GUILD", "ClanFundado", UserList(UserIndex).GuildInfo.ClanFundado)

    For LoopC = 1 To NUMATRIBUTOS
        UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    Next

    For i = 1 To NUMSKILLS
        'str = str & ",SK" & i & "=" & mUser.Stats.UserSkills(i)
        Call WriteVar(UserFile, "SKILLS", "SK" & i, val(UserList(UserIndex).Stats.UserSkills(i)))
    Next i

    Call WriteVar(UserFile, "CONTACTO", "Email", UserList(UserIndex).Email)
    Call WriteVar(UserFile, "INIT", "Genero", val(UserList(UserIndex).Genero))
    Call WriteVar(UserFile, "INIT", "Raza", val(UserList(UserIndex).Raza))
    Call WriteVar(UserFile, "INIT", "Hogar", val(UserList(UserIndex).Hogar))
    Call WriteVar(UserFile, "INIT", "Clase", val(UserList(UserIndex).Clase))
    Call WriteVar(UserFile, "INIT", "Password", UserList(UserIndex).PassWord)
    Call WriteVar(UserFile, "INIT", "Desc", UserList(UserIndex).Desc)

    Call WriteVar(UserFile, "INIT", "Heading", val(UserList(UserIndex).Char.Heading))

    Call WriteVar(UserFile, "INIT", "Head", val(UserList(UserIndex).OrigChar.Head))

    If UserList(UserIndex).flags.Muerto = 0 Then
        Call WriteVar(UserFile, "INIT", "Body", val(UserList(UserIndex).Char.Body))
    End If

    Call WriteVar(UserFile, "INIT", "LastMac", UserList(UserIndex).Mac)
    Call WriteVar(UserFile, "INIT", "LastIP", UserList(UserIndex).ip)
    Call WriteVar(UserFile, "INIT", "Position", UserList(UserIndex).POS.Map & "-" & UserList(UserIndex).POS.X & "-" & UserList(UserIndex).POS.Y)


    Call WriteVar(UserFile, "STATS", "GLD", val(UserList(UserIndex).Stats.GLD))
    Call WriteVar(UserFile, "STATS", "BANCO", val(UserList(UserIndex).Stats.Banco))

    Call WriteVar(UserFile, "STATS", "MET", val(UserList(UserIndex).Stats.MET))
    Call WriteVar(UserFile, "STATS", "MaxHP", val(UserList(UserIndex).Stats.MaxHP))
    Call WriteVar(UserFile, "STATS", "MinHP", val(UserList(UserIndex).Stats.MinHP))

    Call WriteVar(UserFile, "STATS", "FIT", val(UserList(UserIndex).Stats.FIT))
    'Call WriteVar(UserFile, "STATS", "MaxSTA", val(UserList(UserIndex).Stats.MaxSta))
    Call WriteVar(UserFile, "STATS", "MinSTA", val(UserList(UserIndex).Stats.MinSta))


    Call WriteVar(UserFile, "STATS", "MaxMAN", val(UserList(UserIndex).Stats.MaxMAN))
    Call WriteVar(UserFile, "STATS", "MinMAN", val(UserList(UserIndex).Stats.MinMAN))

    Call WriteVar(UserFile, "STATS", "MaxHIT", val(UserList(UserIndex).Stats.MaxHit))
    Call WriteVar(UserFile, "STATS", "MinHIT", val(UserList(UserIndex).Stats.MinHit))

    Call WriteVar(UserFile, "STATS", "MaxAGU", val(UserList(UserIndex).Stats.MaxAGU))
    Call WriteVar(UserFile, "STATS", "MinAGU", val(UserList(UserIndex).Stats.MinAGU))

    Call WriteVar(UserFile, "STATS", "MaxHAM", val(UserList(UserIndex).Stats.MaxHam))
    Call WriteVar(UserFile, "STATS", "MinHAM", val(UserList(UserIndex).Stats.MinHam))

    Call WriteVar(UserFile, "STATS", "SkillPtsLibres", val(UserList(UserIndex).Stats.SkillPts))


    Call WriteVar(UserFile, "STATS", "EXP", val(UserList(UserIndex).Stats.Exp))
    Call WriteVar(UserFile, "STATS", "ELV", val(UserList(UserIndex).Stats.ELV))
    Call WriteVar(UserFile, "STATS", "ELU", val(UserList(UserIndex).Stats.ELU))

    Call WriteVar(UserFile, "MUERTES", "VecesMurioUsuario", val(UserList(UserIndex).Stats.VecesMurioUsuario))
    Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", val(UserList(UserIndex).Stats.NPCsMuertos))

    '[KEVIN]----------------------------------------------------------------------------
    '*******************************************************************************************
    Call WriteVar(UserFile, "BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
    Dim loopd As Integer
    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        Call WriteVar(UserFile, "BancoInventory", "Obj" & loopd, UserList(UserIndex).BancoInvent.Object(loopd).OBJIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount)
    Next loopd
    '*******************************************************************************************
    '[/KEVIN]-----------

    'Save Inv
    Call WriteVar(UserFile, "Inventory", "CantidadItems", val(UserList(UserIndex).Invent.NroItems))

    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(UserIndex).Invent.Object(LoopC).OBJIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).Amount)    '& "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped)
    Next

    Call WriteVar(UserFile, "Inventory", "WEAPONSLOT", val(UserList(UserIndex).Invent.WeaponEqpSlot))
    Call WriteVar(UserFile, "Inventory", "ARMORSLOT", val(UserList(UserIndex).Invent.ArmourEqpSlot))
    Call WriteVar(UserFile, "Inventory", "CASCOSLOT", val(UserList(UserIndex).Invent.CascoEqpSlot))
    Call WriteVar(UserFile, "Inventory", "SHIELDSLOT", val(UserList(UserIndex).Invent.EscudoEqpSlot))
    Call WriteVar(UserFile, "Inventory", "BarcoSlot", val(UserList(UserIndex).Invent.BarcoSlot))
    Call WriteVar(UserFile, "Inventory", "MUNICIONSLOT", val(UserList(UserIndex).Invent.MunicionEqpSlot))
    Call WriteVar(UserFile, "Inventory", "HERRAMIENTASLOT", val(UserList(UserIndex).Invent.HerramientaEqpslot))

    For LoopC = 1 To 3
        Call WriteVar(UserFile, "RECOMPENSAS", "Recompensa" & LoopC, val(UserList(UserIndex).Recompensas(LoopC)))
    Next LoopC

    Dim cad As String

    For LoopC = 1 To MAXUSERHECHIZOS
        cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
        Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, val(cad))
    Next

    Exit Sub

errhandler:
    Call LogError("Error en SaveUser")

End Sub

