Attribute VB_Name = "modHechizos"

Option Explicit
Sub NpcLanzaSpellSobreUser(NpcIndex As Integer, UserIndex As Integer, Spell As Integer)

    If UserList(UserIndex).flags.Privilegios Then Exit Sub

    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub

    Npclist(NpcIndex).CanAttack = 0
    Dim Daño As Integer

    If Hechizos(Spell).SubeHP = 1 Then
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

        Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + Daño
        If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

        Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, Resistencia)
    ElseIf Hechizos(Spell).SubeHP = 2 Then
        Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)

        If Npclist(NpcIndex).MaestroUser = 0 Then Daño = Daño * (1 - UserList(UserIndex).Stats.UserSkills(Resistencia) / 200)

        If UserList(UserIndex).Invent.CascoEqpObjIndex Then
            Dim Obj As ObjData
            Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
            If Obj.Gorro = 1 Then
                Dim absorbido As Integer
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                absorbido = absorbido
                Daño = Maximo(1, Daño - absorbido)
            End If
        End If

        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

        If Not UserList(UserIndex).flags.Quest And UserList(UserIndex).flags.Privilegios = 0 Then
            UserList(UserIndex).Stats.MinHP = Maximo(0, UserList(UserIndex).Stats.MinHP - Daño)
            Call SendUserHP(UserIndex)
        End If

        Call SendData(ToIndex, UserIndex, 0, "%A" & Npclist(NpcIndex).Name & "," & Daño)
        Call SubirSkill(UserIndex, Resistencia)

        If UserList(UserIndex).Stats.MinHP = 0 Then Call UserDie(UserIndex)

    End If

    If Hechizos(Spell).Paraliza > 0 Then
        If UserList(UserIndex).POS.Map = 8 Then
            If UserList(UserIndex).Clase = MAGO Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes usar Inmovilizar en este mapa" & FONTTYPE_FENIX)
            End If
            Exit Sub
        End If
        If UserList(UserIndex).flags.Paralizado = 0 Then
            If UserList(UserIndex).Clase = PIRATA And UserList(UserIndex).Recompensas(3) = 1 Then Exit Sub
            UserList(UserIndex).flags.Paralizado = 1
            UserList(UserIndex).Counters.Paralisis = Timer - 15 * (UserList(UserIndex).Clase = GUERRERO And UserList(UserIndex).Recompensas(3))
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
            Call SendData(ToIndex, UserIndex, 0, ("P9"))
            Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).POS.X & "," & UserList(UserIndex).POS.Y)
        End If
    End If

    If Hechizos(Spell).Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 1
        UserList(UserIndex).Counters.Ceguera = Timer
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        Call SendData(ToIndex, UserIndex, 0, "CEGU")
        Call SendData(ToIndex, UserIndex, 0, "%B")
    End If

    If Hechizos(Spell).RemoverParalisis = 1 Then
        If Npclist(NpcIndex).flags.Paralizado Then
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
        End If
    End If

End Sub
Function TieneHechizo(ByVal i As Integer, UserIndex As Integer) As Boolean

    On Error GoTo errhandler

    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

    Exit Function
errhandler:

End Function
Sub AgregarHechizo(UserIndex As Integer, Slot As Byte)
    Dim hIndex As Integer, j As Integer

    hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex).HechizoIndex

    If Not TieneHechizo(hIndex, UserIndex) Then
        For j = 1 To MAXUSERHECHIZOS
            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next

        If UserList(UserIndex).Stats.UserHechizos(j) Then
            Call SendData(ToIndex, UserIndex, 0, "%C")
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))

            Call QuitarUnItem(UserIndex, CByte(Slot))
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "%D")
    End If

End Sub
Sub Aprenderhechizo(UserIndex As Integer, ByVal hechizoespecial As Integer)
    Dim hIndex As Integer
    Dim j As Integer
    hIndex = hechizoespecial

    If Not TieneHechizo(hIndex, UserIndex) Then

        For j = 1 To MAXUSERHECHIZOS
            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next

        If UserList(UserIndex).Stats.UserHechizos(j) Then
            Call SendData(ToIndex, UserIndex, 0, "%C")
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))

        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "%D")
    End If

End Sub
Sub DecirPalabrasMagicas(ByVal S As String, UserIndex As Integer)
    On Error Resume Next

    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "||" & &HFF00& & "°" & S & "°" & UserList(UserIndex).Char.CharIndex)

End Sub
Function ManaHechizo(UserIndex As Integer, Hechizo As Integer) As Integer

    If UserList(UserIndex).flags.Privilegios > 2 Or UserList(UserIndex).flags.Quest Then Exit Function

    If UserList(UserIndex).Recompensas(3) = 1 And _
       ((UserList(UserIndex).Clase = DRUIDA And Hechizo = 24) Or _
        (UserList(UserIndex).Clase = PALADIN And Hechizo = 10)) Then
        ManaHechizo = 250
    ElseIf UserList(UserIndex).Clase = CLERIGO And UserList(UserIndex).Recompensas(3) = 2 And Hechizo = 11 Then
        ManaHechizo = 1100
    Else: ManaHechizo = Hechizos(Hechizo).ManaRequerido
    End If

End Function
Function PuedeLanzar(UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
    Dim wp2 As WorldPos

    wp2.Map = UserList(UserIndex).flags.TargetMap
    wp2.X = UserList(UserIndex).flags.TargetX
    wp2.Y = UserList(UserIndex).flags.TargetY

    If Not EnPantalla(UserList(UserIndex).POS, wp2, 1) Then Exit Function

    If UserList(UserIndex).flags.Muerto Then
        Call SendData(ToIndex, UserIndex, 0, "MU")
        Exit Function
    End If

    If MapInfo(UserList(UserIndex).POS.Map).NoMagia Then
        Call SendData(ToIndex, UserIndex, 0, "/T")
        Exit Function
    End If

    If UserList(UserIndex).Stats.ELV < Hechizos(HechizoIndex).Nivel Then
        Call SendData(ToIndex, UserIndex, 0, "%%" & Hechizos(HechizoIndex).Nivel)
        Exit Function
    End If

    If UserList(UserIndex).Stats.UserSkills(Magia) < Hechizos(HechizoIndex).MinSkill Then
        Call SendData(ToIndex, UserIndex, 0, "%E")
        Exit Function
    End If

    If UserList(UserIndex).Stats.MinMAN < ManaHechizo(UserIndex, HechizoIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "%F")
        Exit Function
    End If

    If UserList(UserIndex).Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
        Call SendData(ToIndex, UserIndex, 0, "9C")
        Exit Function
    End If

    PuedeLanzar = True

End Function
Sub HechizoInvocacion(UserIndex As Integer, B As Boolean)
    Dim Masc As Integer

    If Not MapInfo(UserList(UserIndex).POS.Map).Pk Then
        Call SendData(ToIndex, UserIndex, 0, "A&")
        Exit Sub
    End If

    If Not UserList(UserIndex).flags.Quest And UserList(UserIndex).NroMascotas >= 3 Then Exit Sub
    If UserList(UserIndex).NroMascotas >= MAXMASCOTAS Then Exit Sub

    Dim H As Integer, j As Integer, ind As Integer, Index As Integer
    Dim TargetPos As WorldPos

    TargetPos.Map = UserList(UserIndex).flags.TargetMap
    TargetPos.X = UserList(UserIndex).flags.TargetX
    TargetPos.Y = UserList(UserIndex).flags.TargetY

    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(j) Then
            If Npclist(UserList(UserIndex).MascotasIndex(j)).Numero = Hechizos(H).NumNPC Then Masc = Masc + 1
        End If
    Next

    If (Hechizos(H).NumNPC = 103 And Masc >= 2 And Not UserList(UserIndex).flags.Quest) Or (Hechizos(H).NumNPC = 94 And Masc >= 1) Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes invocar más mascotas de este tipo." & FONTTYPE_FIGHT)
        Exit Sub
    End If

    For j = 1 To Hechizos(H).Cant
        If (UserList(UserIndex).NroMascotas < 3 Or UserList(UserIndex).flags.Quest) And UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
            ind = SpawnNpc(Hechizos(H).NumNPC, TargetPos, True, False)
            If ind < MAXNPCS Then

                UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1

                Index = FreeMascotaIndex(UserIndex)

                UserList(UserIndex).MascotasIndex(Index) = ind
                UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero

                If UserList(UserIndex).Clase = DRUIDA And UserList(UserIndex).Recompensas(3) = 2 Then
                    If Hechizos(H).NumNPC >= 92 And Hechizos(H).NumNPC <= 94 Then
                        Npclist(ind).Stats.MaxHP = Npclist(ind).Stats.MaxHP + 75
                        Npclist(ind).Stats.MinHP = Npclist(ind).Stats.MaxHP
                    End If
                End If

                If Npclist(ind).Numero = 103 And UserList(UserIndex).Raza <> ELFO_OSCURO Then
                    Npclist(ind).Stats.MaxHP = Npclist(ind).Stats.MaxHP - 200
                    Npclist(ind).Stats.MinHP = Npclist(ind).Stats.MinHP - 200
                End If

                Npclist(ind).MaestroUser = UserIndex
                Npclist(ind).Contadores.TiempoExistencia = Timer
                Npclist(ind).GiveGLD = 0

                Call FollowAmo(ind)
            End If
        Else: Exit For
        End If
    Next

    Call InfoHechizo(UserIndex)
    B = True

End Sub
Sub HechizoTerrenoEstado(UserIndex As Integer, B As Boolean)
    Dim PosCasteada As WorldPos
    Dim TU As Integer
    Dim H As Integer
    Dim i As Integer

    PosCasteada.X = UserList(UserIndex).flags.TargetX
    PosCasteada.Y = UserList(UserIndex).flags.TargetY
    PosCasteada.Map = UserList(UserIndex).flags.TargetMap

    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    If Hechizos(H).Invisibilidad = 2 Then
        For i = 1 To MapInfo(UserList(UserIndex).POS.Map).NumUsers
            TU = MapInfo(UserList(UserIndex).POS.Map).UserIndex(i)
            If EnPantalla(PosCasteada, UserList(TU).POS, -1) And UserList(TU).flags.Invisible = 1 And UserList(TU).flags.AdminInvisible = 0 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(TU).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
            End If
        Next
        B = True
    End If

    Call InfoHechizo(UserIndex)

End Sub
Sub HandleHechizoTerreno(UserIndex As Integer, ByVal uh As Integer)
    Dim B As Boolean

    Select Case Hechizos(uh).Tipo
    Case uInvocacion
        Call HechizoInvocacion(UserIndex, B)
    Case uRadial
        Call HechizoTerrenoEstado(UserIndex, B)
    End Select

    If B Then
        Call SubirSkill(UserIndex, Magia)
        Call QuitarSta(UserIndex, Hechizos(uh).StaRequerido)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizo(UserIndex, uh)
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call SendUserMANASTA(UserIndex)
    End If

End Sub
Sub HandleHechizoUsuario(UserIndex As Integer, ByVal uh As Integer)
    Dim B As Boolean
    Dim tempChr As Integer
    Dim TU, tN As Integer

    tempChr = UserList(UserIndex).flags.TargetUser

    If UserList(tempChr).flags.Protegido = 1 Or UserList(tempChr).flags.Protegido = 2 Then Exit Sub

    Select Case Hechizos(uh).Tipo
    Case uTerreno
        Call HechizoInvocacion(UserIndex, B)
    Case uEstado
        Call HechizoEstadoUsuario(UserIndex, B)
    Case uPropiedades
        Call HechizoPropUsuario(UserIndex, B)
    End Select

    If B Then
        Call SubirSkill(UserIndex, Magia)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizo(UserIndex, uh)
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call QuitarSta(UserIndex, Hechizos(uh).StaRequerido)
        Call SendUserMANASTA(UserIndex)
        Call SendUserHPSTA(UserList(UserIndex).flags.TargetUser)
        UserList(UserIndex).flags.TargetUser = 0
    End If

End Sub
Sub HandleHechizoNPC(UserIndex As Integer, ByVal uh As Integer)
    Dim B As Boolean

    If Npclist(UserList(UserIndex).flags.TargetNpc).flags.NoMagia = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "/U")
        Exit Sub
    End If

    If UserList(UserIndex).flags.Protegido > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar NPC's mientrás estas siendo protegido." & FONTTYPE_FIGHT)
        Exit Sub
    End If

    Select Case Hechizos(uh).Tipo
    Case uEstado
        Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNpc, uh, B, UserIndex)
    Case uPropiedades
        Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNpc, UserIndex, B)
    End Select

    If B Then
        Call SubirSkill(UserIndex, Magia)
        UserList(UserIndex).flags.TargetNpc = 0
        Call QuitarSta(UserIndex, Hechizos(uh).StaRequerido)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizo(UserIndex, uh)
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call SendUserMANASTA(UserIndex)
    End If

End Sub
Sub LanzarHechizo(Index As Integer, UserIndex As Integer)
    Dim uh As Integer
    Dim exito As Boolean

    If UserList(UserIndex).flags.Protegido = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podés tirar hechizos mientras estás siendo protegido por un GM." & FONTTYPE_FIGHT)
        Exit Sub
    ElseIf UserList(UserIndex).flags.Protegido = 2 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podés tirar hechizos tan pronto al conectarte." & FONTTYPE_FIGHT)
        Exit Sub
    End If

    uh = UserList(UserIndex).Stats.UserHechizos(Index)

    If (UserList(UserIndex).POS.Map = 148 Or UserList(UserIndex).POS.Map = 150) And (Hechizos(uh).Invoca > 0 Or Hechizos(uh).SubeHP = 2 Or Hechizos(uh).Invisibilidad = 1 Or Hechizos(uh).Paraliza > 0 Or Hechizos(uh).Estupidez = 1) Then
        Call SendData(ToIndex, UserIndex, 0, "||Una extraña energía te impide lanzar este hechizo..." & FONTTYPE_INFO)
        Exit Sub
    End If

    If TiempoTranscurrido(UserList(UserIndex).Counters.LastHechizo) < IntervaloUserPuedeCastear Then Exit Sub
    If TiempoTranscurrido(UserList(UserIndex).Counters.LastGolpe) < IntervaloUserPuedeGolpeHechi Then Exit Sub
    UserList(UserIndex).Counters.LastHechizo = Timer
    Call SendData(ToIndex, UserIndex, 0, "LH")

    If Hechizos(uh).Baculo > 0 And (UserList(UserIndex).Clase = DRUIDA Or UserList(UserIndex).Clase = MAGO Or UserList(UserIndex).Clase = NIGROMANTE) Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo < Hechizos(uh).Baculo Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "BN")
            Else: Call SendData(ToIndex, UserIndex, 0, "||Debes equiparte un báculo de mayor rango para lanzar este hechizo." & FONTTYPE_INFO)
            End If
            Exit Sub
        End If
    End If

    If PuedeLanzar(UserIndex, uh) Then
        Select Case Hechizos(uh).Target

        Case uUsuarios
            If UserList(UserIndex).flags.TargetUser Then
                If UserList(UserList(UserIndex).flags.TargetUser).POS.Y - UserList(UserIndex).POS.Y >= 7 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos para lanzar ese hechizo." & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call HandleHechizoUsuario(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
            End If

        Case uNPC
            If UserList(UserIndex).flags.TargetNpc Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If

        Case uUsuariosYnpc
            If UserList(UserIndex).flags.TargetUser Then
                If UserList(UserList(UserIndex).flags.TargetUser).POS.Y - UserList(UserIndex).POS.Y >= 7 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos para lanzar ese hechizo." & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call HandleHechizoUsuario(UserIndex, uh)
            ElseIf UserList(UserIndex).flags.TargetNpc Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If

        Case uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)

        Case uArea
            Call HandleHechizoArea(UserIndex, uh)

        End Select
    End If

End Sub
Sub HandleHechizoArea(UserIndex As Integer, ByVal uh As Integer)
    On Error GoTo Error
    Dim TargetPos As WorldPos
    Dim X2 As Integer, Y2 As Integer
    Dim UI As Integer
    Dim B As Boolean

    TargetPos.Map = UserList(UserIndex).flags.TargetMap
    TargetPos.X = UserList(UserIndex).flags.TargetX
    TargetPos.Y = UserList(UserIndex).flags.TargetY

    For X2 = TargetPos.X - Hechizos(uh).RadioX To TargetPos.X + Hechizos(uh).RadioX
        For Y2 = TargetPos.Y - Hechizos(uh).RadioY To TargetPos.Y + Hechizos(uh).RadioY
            UI = MapData(TargetPos.Map, X2, Y2).UserIndex
            If UI > 0 Then
                UserList(UserIndex).flags.TargetUser = UI
                Select Case Hechizos(uh).Tipo
                Case uEstado
                    Call HechizoEstadoUsuario(UserIndex, B)
                Case uPropiedades
                    Call HechizoPropUsuario(UserIndex, B)
                End Select
            End If
        Next
    Next

    If B Then
        Call SubirSkill(UserIndex, Magia)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizo(UserIndex, uh)
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call QuitarSta(UserIndex, Hechizos(uh).StaRequerido)
        Call SendUserMANASTA(UserIndex)
        UserList(UserIndex).flags.TargetUser = 0
    End If

    Exit Sub
Error:
    Call LogError("Error en HandleHechizoArea")
End Sub
Public Function Amigos(UserIndex As Integer, UI As Integer) As Boolean

    Amigos = (((UserList(UserIndex).Faccion.Bando = UserList(UI).Faccion.Bando) Or (EsNewbie(UI)) Or (EsNewbie(UserIndex)))) Or (UserList(UserIndex).POS.Map = 8) Or (UserList(UserIndex).Faccion.Bando = Neutral)

End Function
Sub HechizoEstadoUsuario(UserIndex As Integer, B As Boolean)
    Dim H As Integer, TU As Integer, HechizoBueno As Boolean

    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    TU = UserList(UserIndex).flags.TargetUser

    HechizoBueno = Hechizos(H).RemoverParalisis Or Hechizos(H).CuraVeneno Or Hechizos(H).Invisibilidad Or Hechizos(H).Revivir Or Hechizos(H).Flecha Or Hechizos(H).Estupidez = 2 Or Hechizos(H).Transforma

    If HechizoBueno Then
        If Not Amigos(UserIndex, TU) Then
            Call SendData(ToIndex, UserIndex, 0, "2F")
            Exit Sub
        End If
    Else
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        'If UserList(UserIndex).flags.Invisible Then Call BajarInvisible(UserIndex)
        Call UsuarioAtacadoPorUsuario(UserIndex, TU)
    End If

    If Hechizos(H).Envenena Then
        UserList(TU).flags.Envenenado = Hechizos(H).Envenena
        UserList(TU).flags.EstasEnvenenado = Timer
        UserList(TU).Counters.Veneno = Timer
        Call InfoHechizo(UserIndex)
        B = True
        Exit Sub
    End If

    If Hechizos(H).Maldicion = 1 Then
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        B = True
        Exit Sub
    End If

    If Hechizos(H).Paraliza > 0 Then
        If UserList(UserIndex).POS.Map = 8 And UserList(UserIndex).Clase = MAGO Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes usar Inmovilizar en este mapa" & FONTTYPE_FENIX)
            Exit Sub
        End If
        If UserList(TU).flags.Paralizado = 0 Then
            If (UserList(TU).Clase = MINERO And UserList(TU).Recompensas(2) = 1) Or (UserList(TU).Clase = PIRATA And UserList(TU).Recompensas(3) = 1) Then
                Call SendData(ToIndex, UserIndex, 0, "%&")
                Exit Sub
            End If

            UserList(TU).flags.QuienParalizo = UserIndex
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = Timer - 15 * Buleano(UserList(TU).Clase = GUERRERO And UserList(TU).Recompensas(3) = 2)
            Call SendData(ToIndex, TU, 0, "PU" & UserList(TU).POS.X & "," & UserList(TU).POS.Y)
            Call SendData(ToIndex, TU, 0, ("P9"))
            Call InfoHechizo(UserIndex)
            B = True
            Exit Sub
        End If
    End If

    If Hechizos(H).Ceguera = 1 Then
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = Timer
        Call SendData(ToIndex, TU, 0, "CEGU")
        Call InfoHechizo(UserIndex)
        B = True
        Exit Sub
    End If

    If Hechizos(H).Estupidez = 1 Then
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Estupidez = Timer
        Call SendData(ToIndex, TU, 0, "DUMB")
        Call InfoHechizo(UserIndex)
        B = True
        Exit Sub
    End If

    If Hechizos(H).Transforma = 1 Then
        If UserList(TU).flags.Transformado = 0 Then
            If UserList(TU).Stats.ELV > 39 And UserList(TU).Raza = ELFO And UserList(TU).Clase = DRUIDA Then
                Call DoMetamorfosis(UserIndex)
            Else
                Call SendData(ToIndex, UserIndex, 0, "{E")
            End If
            Call InfoHechizo(UserIndex)
            B = True
            Exit Sub
        End If
    End If

    If Hechizos(H).Revivir = 1 Then
        If UserList(TU).flags.Muerto Then
            Call RevivirUsuario(UserIndex, TU, UserList(UserIndex).Clase = CLERIGO And UserList(UserIndex).Recompensas(3) = 2)
            Call InfoHechizo(UserIndex)
            B = True
            Exit Sub
        End If
    End If

    If UserList(TU).flags.Muerto Then
        Call SendData(ToIndex, UserIndex, 0, "8C")
        Exit Sub
    End If

    If Hechizos(H).Estupidez = 2 Then
        If UserList(TU).flags.Estupidez = 1 Then
            UserList(TU).flags.Estupidez = 0
            UserList(TU).Counters.Estupidez = 0
            Call SendData(ToIndex, TU, 0, "NESTUP")
            Call InfoHechizo(UserIndex)
            B = True
            Exit Sub
        End If
    End If

    If Hechizos(H).Flecha = 1 Then
        If TU <> UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "||Este hechizo solo puedes usarlo sobre ti mismo." & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(TU).flags.BonusFlecha = True
        UserList(TU).Counters.BonusFlecha = Timer
        Call InfoHechizo(UserIndex)
        B = True
        Exit Sub
    End If

    If Hechizos(H).RemoverParalisis = 1 Then
        If UserList(TU).flags.Paralizado Then
            Call SendData(ToIndex, TU, 0, "P8")
            UserList(TU).flags.Paralizado = 0
            UserList(TU).flags.QuienParalizo = 0
            Call InfoHechizo(UserIndex)
            B = True
            Exit Sub
        End If
    End If

    If Hechizos(H).Invisibilidad = 1 Then
        If UserList(UserIndex).POS.Map = 8 Or UserList(UserIndex).POS.Map = 15 Or UserList(UserIndex).POS.Map = 16 Or UserList(UserIndex).POS.Map = 10 Or UserList(UserIndex).POS.Map = 11 Or UserList(UserIndex).POS.Map = 12 Or UserList(UserIndex).POS.Map = 14 Or UserList(UserIndex).POS.Map = 13 Or UserList(UserIndex).POS.Map = 88 Or UserList(UserIndex).POS.Map = 7 Or UserList(UserIndex).POS.Map = 17 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes hacerte invisible en este mapa" & FONTTYPE_FENIX)
            Exit Sub
        End If
        If UserList(TU).flags.Invisible Then Exit Sub
        UserList(TU).flags.Invisible = 1
        UserList(TU).Counters.Invisibilidad = Timer
        Call SendData(ToMap, 0, UserList(TU).POS.Map, ("V3" & UserList(TU).Char.CharIndex & ",1"))
        Call InfoHechizo(UserIndex)
        B = True
        Exit Sub
    End If

    If Hechizos(H).CuraVeneno = 1 Then
        If UserList(TU).flags.Envenenado = 1 Then
            UserList(TU).flags.Envenenado = 0
            Call InfoHechizo(UserIndex)
            B = True
            Exit Sub
        Else
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no está envenenado." & FONTTYPE_FIGHT)
            Exit Sub
        End If
    End If

    If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        B = True
        Exit Sub
    End If

    If Hechizos(H).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        B = True
        Exit Sub
    End If

End Sub
Sub HechizoEstadoNPC(NpcIndex As Integer, ByVal hIndex As Integer, B As Boolean, UserIndex As Integer)

    If Npclist(NpcIndex).Attackable = 0 Then Exit Sub

    If Hechizos(hIndex).Invisibilidad = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Invisible = 1
        B = True
    End If

    If Hechizos(hIndex).Envenena = 1 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "NO")
            Exit Sub
        End If
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Envenenado = 1
        B = True
    End If

    If Hechizos(hIndex).CuraVeneno = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Envenenado = 0
        B = True
    End If

    If Hechizos(hIndex).Maldicion = 1 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "NO")
            Exit Sub
        End If
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Maldicion = 1
        B = True
    End If

    If Hechizos(hIndex).RemoverMaldicion = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Maldicion = 0
        B = True
    End If

    If Hechizos(hIndex).Bendicion = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Bendicion = 1
        B = True
    End If

    If Hechizos(hIndex).Paraliza Then
        If UserList(UserIndex).POS.Map = 8 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes usar Inmovilizar en este mapa" & FONTTYPE_FENIX)
            Exit Sub
        End If

        If Npclist(NpcIndex).flags.QuienParalizo <> 0 And Npclist(NpcIndex).flags.QuienParalizo <> UserIndex Then Exit Sub
        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = Hechizos(hIndex).Paraliza
            Npclist(NpcIndex).flags.QuienParalizo = UserIndex
            If Npclist(NpcIndex).flags.PocaParalisis = 1 Then
                Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado / 4
            Else: Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            End If
            B = True
        Else: Call SendData(ToIndex, UserIndex, 0, "7D")
        End If
    End If

    If Hechizos(hIndex).RemoverParalisis = 1 Then
        If Npclist(NpcIndex).flags.QuienParalizo = UserIndex Or Npclist(NpcIndex).MaestroUser = UserIndex Then
            If Npclist(NpcIndex).flags.Paralizado Then
                Call InfoHechizo(UserIndex)
                Npclist(NpcIndex).flags.Paralizado = 0
                Npclist(NpcIndex).Contadores.Paralisis = 0
                Npclist(NpcIndex).flags.QuienParalizo = 0
                B = True
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "8D")
        End If
    End If

End Sub
Sub VerNPCMuere(ByVal NpcIndex As Integer, ByVal Daño As Long, ByVal UserIndex As Integer)

    If Npclist(NpcIndex).AutoCurar = 0 Then Npclist(NpcIndex).Stats.MinHP = Maximo(0, Npclist(NpcIndex).Stats.MinHP - Daño)

    If Npclist(NpcIndex).Stats.MinHP <= 0 Then
        If Npclist(NpcIndex).flags.Snd3 Then Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).POS.Map, "TW" & Npclist(NpcIndex).flags.Snd3)

        If UserIndex Then
            If UserList(UserIndex).NroMascotas Then
                Dim T As Integer
                For T = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
                    If UserList(UserIndex).MascotasIndex(T) Then
                        If Npclist(UserList(UserIndex).MascotasIndex(T)).TargetNpc = NpcIndex Then Call FollowAmo(UserList(UserIndex).MascotasIndex(T))
                    End If
                Next
            End If
            Call AddtoVar(UserList(UserIndex).Stats.NPCsMuertos, 1, 32000)

            UserList(UserIndex).flags.TargetNpc = 0
            UserList(UserIndex).flags.TargetNpcTipo = 0
        End If

        Call MuereNpc(NpcIndex, UserIndex)
    End If

End Sub
Sub ExperienciaPorGolpe(UserIndex As Integer, ByVal NpcIndex As Integer, Daño As Integer)
    Dim ExpDada As Long

    Daño = Minimo(Daño, Npclist(NpcIndex).Stats.MinHP)

    ExpDada = Npclist(NpcIndex).GiveEXP * (Daño / Npclist(NpcIndex).Stats.MaxHP) / 2

    If Daño >= Npclist(NpcIndex).Stats.MinHP Then ExpDada = ExpDada + Npclist(NpcIndex).GiveEXP / 2
    If ModoQuest Then ExpDada = ExpDada / 2

    If UserList(UserIndex).flags.Party = 0 Then
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpDada
        If Daño >= Npclist(NpcIndex).Stats.MinHP Then
            Call SendData(ToIndex, UserIndex, 0, "EL" & ExpDada)
        Else: Call SendData(ToIndex, UserIndex, 0, "EX" & ExpDada)
        End If
        Call SendUserEXP(UserIndex)
        Call CheckUserLevel(UserIndex)
        Exit Sub
    Else: Call RepartirExp(UserIndex, ExpDada, Daño >= Npclist(NpcIndex).Stats.MinHP)
    End If

End Sub
Sub HechizoPropNPC(ByVal hIndex As Integer, NpcIndex As Integer, UserIndex As Integer, B As Boolean)
    Dim Daño As Integer

    If Npclist(NpcIndex).Attackable = 0 Then Exit Sub

    If Hechizos(hIndex).SubeHP = 1 Then
        Daño = DañoHechizo(UserIndex, hIndex)

        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "||" & vbGreen & "°+" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))

        Call InfoHechizo(UserIndex)
        Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, Daño, Npclist(NpcIndex).Stats.MaxHP)
        Call SendData(ToIndex, UserIndex, 0, "CU" & Daño)
        B = True
    ElseIf Hechizos(hIndex).SubeHP = 2 Then

        Daño = DañoHechizo(UserIndex, hIndex)

        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name = "Báculo Oscuro" Then
            Daño = 1.05 * Daño
        ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo = Hechizos(hIndex).Baculo Then
            Daño = 0.95 * Daño
        End If

        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "NO")
            Exit Sub
        End If

        If UserList(UserIndex).Faccion.Bando <> Neutral And Npclist(NpcIndex).MaestroUser Then
            If Not PuedeAtacarMascota(UserIndex, (Npclist(NpcIndex).MaestroUser)) Then Exit Sub
        End If

        If UserList(UserIndex).Faccion.Bando <> Neutral And UserList(UserIndex).Faccion.Bando = Npclist(NpcIndex).flags.Faccion Then
            Call SendData(ToIndex, UserIndex, 0, Mensajes(Npclist(NpcIndex).flags.Faccion, 19))
            Exit Sub
        End If

        Call InfoHechizo(UserIndex)
        B = True
        Call NpcAtacado(NpcIndex, UserIndex)

        If Npclist(NpcIndex).flags.Snd2 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & Npclist(NpcIndex).flags.Snd2)

        Call SendData(ToIndex, UserIndex, 0, "X2" & Daño)

        Call ExperienciaPorGolpe(UserIndex, NpcIndex, Daño)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "||" & vbRed & "°-" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))

        Call VerNPCMuere(NpcIndex, Daño, UserIndex)
    End If

End Sub
Sub InfoHechizo(UserIndex As Integer)
    Dim H As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & Hechizos(H).WAV)

    If UserList(UserIndex).flags.TargetUser Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    ElseIf UserList(UserIndex).flags.TargetNpc Then
        Call SendData(ToPCArea, UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).POS.Map, "CFX" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    End If

    If UserList(UserIndex).flags.TargetUser Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name & FONTTYPE_ATACO)
            Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).flags.TargetNpc Then
        Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_ATACO)
    End If

End Sub
Function DañoHechizo(UserIndex As Integer, Hechizo As Integer) As Integer

    DañoHechizo = RandomNumber(Hechizos(Hechizo).MinHP + 5 * Buleano(UserList(UserIndex).Clase = BARDO And UserList(UserIndex).Recompensas(3) = 2 And (Hechizo = 23 Or Hechizo = 25)) _
                             + 10 * Buleano(UserList(UserIndex).Clase = NIGROMANTE And UserList(UserIndex).Recompensas(3) = 1) _
                             + 20 * Buleano(UserList(UserIndex).Clase = CLERIGO And UserList(UserIndex).Recompensas(3) = 1 And Hechizo = 5) _
                             + 10 * Buleano(UserList(UserIndex).Clase = MAGO And UserList(UserIndex).Recompensas(3) = 2 And Hechizo = 25), _
                               Hechizos(Hechizo).MaxHP + 5 * Buleano(UserList(UserIndex).Clase = BARDO And UserList(UserIndex).Recompensas(3) = 2 And (Hechizo = 23 Or Hechizo = 25)) _
                             + 20 * Buleano(UserList(UserIndex).Clase = CLERIGO And UserList(UserIndex).Recompensas(3) = 1 And Hechizo = 5) _
                             + 10 * Buleano(UserList(UserIndex).Clase = MAGO And UserList(UserIndex).Recompensas(3) = 1 And Hechizo = 23))

    DañoHechizo = DañoHechizo + Porcentaje(DañoHechizo, 3 * UserList(UserIndex).Stats.ELV)

End Function
Sub HechizoPropUsuario(UserIndex As Integer, B As Boolean)
    Dim H As Integer
    Dim Daño As Integer
    Dim tempChr As Integer
    Dim reducido As Integer
    Dim HechizoBueno As Boolean
    Dim msg As String

    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tempChr = UserList(UserIndex).flags.TargetUser

    HechizoBueno = Hechizos(H).SubeHam = 1 Or Hechizos(H).SubeSed = 1 Or Hechizos(H).SubeHP = 1 Or Hechizos(H).SubeAgilidad = 1 Or Hechizos(H).SubeFuerza = 1 Or Hechizos(H).SubeFuerza = 3 Or Hechizos(H).SubeMana = 1 Or Hechizos(H).SubeSta = 1

    If HechizoBueno And Not Amigos(UserIndex, tempChr) Then
        Call SendData(ToIndex, UserIndex, 0, "2F")
        Exit Sub
    ElseIf Not HechizoBueno Then
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        'If UserList(UserIndex).flags.Invisible Then Call BajarInvisible(UserIndex)
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If

    If Hechizos(H).Revivir = 0 And UserList(tempChr).flags.Muerto Then Exit Sub

    If Hechizos(H).SubeHam = 1 Then

        Daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
        Call InfoHechizo(UserIndex)

        Call AddtoVar(UserList(tempChr).Stats.MinHam, Daño, UserList(tempChr).Stats.MaxHam)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
        End If

        Call EnviarHyS(tempChr)
        B = True

    ElseIf Hechizos(H).SubeHam = 2 Then
        Daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
        UserList(tempChr).Stats.MinHam = Maximo(0, UserList(tempChr).Stats.MinHam - Daño)

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
        End If
        If UserList(tempChr).Stats.MinHam = 0 Then UserList(tempChr).flags.Hambre = 1
        Call EnviarHyS(tempChr)
        B = True
    End If


    If Hechizos(H).SubeSed = 1 Then

        Call AddtoVar(UserList(tempChr).Stats.MinAGU, Daño, UserList(tempChr).Stats.MaxAGU)
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
        End If

        B = True

    ElseIf Hechizos(H).SubeSed = 2 Then
        Daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
        UserList(tempChr).Stats.MinAGU = Maximo(0, UserList(tempChr).Stats.MinAGU - Daño)
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
        End If

        If UserList(tempChr).Stats.MinAGU = 0 Then UserList(tempChr).flags.Sed = 1
        B = True
    ElseIf Hechizos(H).SubeSed = 3 Then

        UserList(tempChr).Stats.MinAGU = 0
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).Stats.MinSta = 0
        UserList(tempChr).flags.Sed = 1
        UserList(tempChr).flags.Hambre = 1

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "S3" & UserList(tempChr).Name)
            Call SendData(ToIndex, tempChr, 0, "S4" & UserList(UserIndex).Name)
        Else
            Call SendData(ToIndex, UserIndex, 0, "S5")
        End If
        Call SendData(ToIndex, tempChr, 0, "2G")

        B = True
    End If


    If Hechizos(H).SubeAgilidad = 1 Then
        Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)

        UserList(tempChr).flags.DuracionEfecto = Timer
        Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2, MAXATRIBUTOS))
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        UserList(tempChr).flags.TomoPocion = True
        B = True

    ElseIf Hechizos(H).SubeAgilidad = 2 Then
        UserList(tempChr).flags.TomoPocion = True
        Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call RestVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, MINATRIBUTOS)
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        B = True
    ElseIf Hechizos(H).SubeAgilidad = 3 Then
        UserList(tempChr).flags.TomoPocion = True
        Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call RestVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, MINATRIBUTOS)
        Call RestVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, MINATRIBUTOS)
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        B = True
    End If


    If Hechizos(H).SubeFuerza = 1 Then
        Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call AddtoVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(fuerza) * 2, MAXATRIBUTOS))
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        UserList(tempChr).flags.TomoPocion = True
        B = True
    ElseIf Hechizos(H).SubeFuerza = 2 Then
        UserList(tempChr).flags.TomoPocion = True
        Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call RestVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, MINATRIBUTOS)
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        B = True
    ElseIf Hechizos(H).SubeFuerza = 3 Then
        Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call AddtoVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(fuerza) * 2, MAXATRIBUTOS))
        Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2, MAXATRIBUTOS))
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        UserList(tempChr).flags.TomoPocion = True
        B = True
    End If


    If Hechizos(H).SubeHP = 1 Then
        If UserList(tempChr).flags.Muerto = 1 Then Exit Sub

        If UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP Then
            Call SendData(ToIndex, UserIndex, 0, "9D")
            Exit Sub
        End If
        Daño = DañoHechizo(UserIndex, H)

        Call AddtoVar(UserList(tempChr).Stats.MinHP, Daño, UserList(tempChr).Stats.MaxHP)
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "R3" & Daño & "," & UserList(tempChr).Name)
            Call SendData(ToIndex, tempChr, 0, "R4" & UserList(UserIndex).Name & "," & Daño)
        Else
            Call SendData(ToIndex, UserIndex, 0, "R5" & Daño)
        End If
        B = True
    ElseIf Hechizos(H).SubeHP = 2 Then
        Daño = DañoHechizo(UserIndex, H)

        If Hechizos(H).Baculo > 0 And (UserList(UserIndex).Clase = DRUIDA Or UserList(UserIndex).Clase = MAGO Or UserList(UserIndex).Clase = NIGROMANTE) Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo < Hechizos(H).Baculo Then
                Call SendData(ToIndex, UserIndex, 0, "BN")
                Exit Sub
            Else
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name = "Báculo Oscuro" Then
                    Daño = 1.05 * Daño
                ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo = Hechizos(H).Baculo Then
                    Daño = 0.95 * Daño
                End If
            End If
        End If

        If UserList(tempChr).Invent.CascoEqpObjIndex Then
            Dim Obj As ObjData
            Obj = ObjData(UserList(tempChr).Invent.CascoEqpObjIndex)
            If Obj.Gorro = 1 Then Daño = Maximo(1, (1 - (RandomNumber(Obj.MinDef, Obj.MaxDef) / 100)) * Daño)
            Daño = Maximo(1, Daño)
        End If

        If Not UserList(tempChr).flags.Quest Then UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - Daño
        Call InfoHechizo(UserIndex)

        Call SendData(ToIndex, UserIndex, 0, "6B" & Daño & "," & UserList(tempChr).Name)
        Call SendData(ToIndex, tempChr, 0, "7B" & Daño & "," & UserList(UserIndex).Name)

        If UserList(tempChr).Stats.MinHP > 0 Then
            Call SubirSkill(tempChr, Resistencia)
        Else
            Call ContarMuerte(tempChr, UserIndex)
            UserList(tempChr).Stats.MinHP = 0
            Call ActStats(tempChr, UserIndex)
        End If

        B = True
    End If


    If Hechizos(H).SubeMana = 1 Then
        Call AddtoVar(UserList(tempChr).Stats.MinMAN, Daño, UserList(tempChr).Stats.MaxMAN)
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
        End If

        B = True

    ElseIf Hechizos(H).SubeMana = 2 Then

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
        End If

        UserList(tempChr).Stats.MinMAN = Maximo(0, UserList(tempChr).Stats.MinMAN - Daño)
        B = True

    End If


    If Hechizos(H).SubeSta = 1 Then
        Call AddtoVar(UserList(tempChr).Stats.MinSta, Daño, UserList(tempChr).Stats.MaxSta)

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
        End If
        B = True
    ElseIf Hechizos(H).SubeSta = 2 Then
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
        End If
        Call QuitarSta(tempChr, Daño)
        B = True
    End If

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, UserIndex As Integer, Slot As Byte)
    Dim LoopC As Byte

    If Not UpdateAll Then
        If UserList(UserIndex).Stats.UserHechizos(Slot) Then
            Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
        Else
            Call ChangeUserHechizo(UserIndex, Slot, 0)
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "6H")
        For LoopC = 1 To MAXUSERHECHIZOS
            If UserList(UserIndex).Stats.UserHechizos(LoopC) Then
                Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
            End If
        Next
    End If

End Sub
Sub ChangeUserHechizo(UserIndex As Integer, Slot As Byte, ByVal Hechizo As Integer)

    UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo

    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
        Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)
    Else
        Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & "0" & "," & "Nada")
    End If

End Sub
Public Sub DesplazarHechizo(UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Byte)

    If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
    If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

    Dim TempHechizo As Integer

    If Dire = 1 Then
        If CualHechizo = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "%G")
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

            Call UpdateUserHechizos(False, UserIndex, CualHechizo - 1)
        End If
    Else
        If CualHechizo = MAXUSERHECHIZOS Then
            Call SendData(ToIndex, UserIndex, 0, "%G")
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

            Call UpdateUserHechizos(False, UserIndex, CualHechizo + 1)
        End If
    End If

    Call UpdateUserHechizos(False, UserIndex, CualHechizo)

End Sub

