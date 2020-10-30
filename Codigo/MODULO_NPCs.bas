Attribute VB_Name = "NPCs"
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

Option Explicit
Public rdata As String
Sub QuitarMascota(UserIndex As Integer, ByVal NpcIndex As Integer)
    Dim i As Integer

    UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1

    For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
            UserList(UserIndex).MascotasIndex(i) = 0
            UserList(UserIndex).MascotasType(i) = 0
            Exit For
        End If
    Next

End Sub
Sub QuitarMascotaNpc(Maestro As Integer, ByVal Mascota As Integer)

    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

End Sub
Sub MuereNpc(ByVal NpcIndex As Integer, UserIndex As Integer)
    On Error GoTo errhandler
    Dim Exp As Long
    Dim MiNPC As Npc
    MiNPC = Npclist(NpcIndex)

    Call QuitarNPC(NpcIndex)

    If MiNPC.MaestroUser = 0 Then
        If UserIndex Then Call NPCTirarOro(MiNPC, UserIndex)
        If UserIndex Then Call NPC_TIRAR_ITEMS(MiNPC, UserIndex)
    End If

    If UserIndex > 0 Then Call SubirSkill(UserIndex, Supervivencia, 40)
    Call ReSpawnNpc(MiNPC)

    Exit Sub

errhandler:
    Call LogError("Error en MuereNpc " & Err.Description)

End Sub
Function NPCListable(NpcIndex As Integer) As Boolean

    NPCListable = (Npclist(NpcIndex).Attackable And Not Npclist(NpcIndex).flags.Respawn)

End Function
Sub QuitarNPC(ByVal NpcIndex As Integer)
    On Error GoTo errhandler
    Dim i As Integer

    Npclist(NpcIndex).flags.NPCActive = False

    If NPCListable(NpcIndex) Then Call QuitarNPCDeLista(Npclist(NpcIndex).Numero, Npclist(NpcIndex).POS.Map)

    Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).POS.Map, "QDL" & Npclist(NpcIndex).Char.CharIndex)

    If InMapBounds(Npclist(NpcIndex).POS.X, Npclist(NpcIndex).POS.Y) Then Call EraseNPCChar(ToMap, 0, Npclist(NpcIndex).POS.Map, NpcIndex)

    If Npclist(NpcIndex).MaestroUser Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
    If Npclist(NpcIndex).MaestroNpc Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)

    Npclist(NpcIndex) = NpcNoIniciado

    For i = LastNPC To 1 Step -1
        If Npclist(i).flags.NPCActive Then
            LastNPC = i
            Exit For
        End If
    Next

    If NumNPCs Then NumNPCs = NumNPCs - 1

    Exit Sub

errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC-" & Err.Description)

End Sub
Function TestSpawnTrigger(POS As WorldPos) As Boolean

    If Not InMapBounds(POS.X, POS.Y) Or Not MapaValido(POS.Map) Then Exit Function

    TestSpawnTrigger = _
    MapData(POS.Map, POS.X, POS.Y).trigger <> 3 And _
                       MapData(POS.Map, POS.X, POS.Y).trigger <> 2 And _
                       MapData(POS.Map, POS.X, POS.Y).trigger <> 1

End Function
Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)


    Dim POS As WorldPos
    Dim newpos As WorldPos
    Dim nIndex As Integer
    Dim PosicionValida As Boolean
    Dim Iteraciones As Long
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    On Error GoTo Error

    nIndex = OpenNPC(NroNPC)

    If nIndex > MAXNPCS Then Exit Sub


    If InMapBounds(OrigPos.X, OrigPos.Y) Then

        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).POS = OrigPos

    Else

        POS.Map = mapa

        Do While Not PosicionValida
            DoEvents

            POS.X = CInt(Rnd * 100 + 1)
            POS.Y = CInt(Rnd * 100 + 1)

            Call ClosestLegalPos(POS, newpos, Npclist(nIndex).flags.AguaValida = 1)


            If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida = 1) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then

                Npclist(nIndex).POS.Map = newpos.Map
                Npclist(nIndex).POS.X = newpos.X
                Npclist(nIndex).POS.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0

            End If


            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                Call QuitarNPC(nIndex)
                Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                Exit Sub
            End If
        Loop


        Map = newpos.Map
        X = Npclist(nIndex).POS.X
        Y = Npclist(nIndex).POS.Y
    End If


    Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)

    If NPCListable(nIndex) Then Call AgregarNPC(Npclist(nIndex).Numero, mapa)
    Exit Sub
Error:

    Call LogError("Error en CrearNPC." & Map & " " & X & " " & Y & " " & nIndex)
End Sub
Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer, Map As Integer, X As Integer, Y As Integer)
    Dim CharIndex As Integer

    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If

    MapData(Map, X, Y).NpcIndex = NpcIndex

    Call SendData(sndRoute, sndIndex, sndMap, ("CC" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & X & "," & Y))

End Sub

Sub ChangeNPCChar(NpcIndex As Integer, Body As Integer, Head As Integer, ByVal Heading As Byte)

    If Npclist(NpcIndex).Char.Body = Body And _
       Npclist(NpcIndex).Char.Head = Head And _
       Npclist(NpcIndex).Char.Heading = Heading Then Exit Sub
    If NpcIndex Then
        Npclist(NpcIndex).Char.Body = Body
        Npclist(NpcIndex).Char.Head = Head
        Npclist(NpcIndex).Char.Heading = Heading
        Call SendData(ToNPCAreaG, NpcIndex, Npclist(NpcIndex).POS.Map, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
    End If

End Sub

Sub EraseNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer)

    If Npclist(NpcIndex).Char.CharIndex Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

    If Npclist(NpcIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar < 1 Then Exit Do
        Loop
    End If


    MapData(Npclist(NpcIndex).POS.Map, Npclist(NpcIndex).POS.X, Npclist(NpcIndex).POS.Y).NpcIndex = 0


    Call SendData(ToMap, 0, Npclist(NpcIndex).POS.Map, "BP" & Npclist(NpcIndex).Char.CharIndex)


    Npclist(NpcIndex).Char.CharIndex = 0



    NumChars = NumChars - 1


End Sub
Sub MoveNPCChar(NpcIndex As Integer, ByVal nHeading As Byte)
    On Error GoTo errh
    Dim nPos As WorldPos

    If Npclist(NpcIndex).AutoCurar = 1 Then Exit Sub

    nPos = Npclist(NpcIndex).POS
    Call HeadtoPos(nHeading, nPos)

    If (Npclist(NpcIndex).MaestroUser And LegalPos(Npclist(NpcIndex).POS.Map, nPos.X, nPos.Y)) Or LegalPosNPC(Npclist(NpcIndex).POS.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida = 1) Then
        If (Npclist(NpcIndex).flags.AguaValida = 0 And MapData(Npclist(NpcIndex).POS.Map, nPos.X, nPos.Y).Agua = 1) Or (Npclist(NpcIndex).flags.TierraInvalida = 1 And MapData(Npclist(NpcIndex).POS.Map, nPos.X, nPos.Y).Agua = 0) Then Exit Sub

        Call SendData(ToNPCAreaG, NpcIndex, Npclist(NpcIndex).POS.Map, "MP" & (Npclist(NpcIndex).Char.CharIndex) & "," & (nPos.X) & "," & (nPos.Y))


        MapData(Npclist(NpcIndex).POS.Map, Npclist(NpcIndex).POS.X, Npclist(NpcIndex).POS.Y).NpcIndex = 0
        Npclist(NpcIndex).POS = nPos
        Npclist(NpcIndex).Char.Heading = nHeading
        MapData(Npclist(NpcIndex).POS.Map, Npclist(NpcIndex).POS.X, Npclist(NpcIndex).POS.Y).NpcIndex = NpcIndex
    Else
        If Npclist(NpcIndex).Movement = NPC_PATHFINDING Then Npclist(NpcIndex).PFINFO.PathLenght = 0
    End If

    Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)

End Sub
Function Bin(N)

    Dim S As String, i As Integer, uu, T

    uu = Int(Log(N) / Log(2))

    For i = 0 To uu
        S = (N Mod 2) & S
        T = N / 2
        N = Int(T)
    Next
    Bin = S

End Function
Function NextOpenNPC() As Integer
    On Error GoTo errhandler

    Dim LoopC As Integer

    For LoopC = 1 To MAXNPCS + 1
        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next

    NextOpenNPC = LoopC

    Exit Function
errhandler:
    Call LogError("Error en NextOpenNPC")
End Function
Sub NpcEnvenenarUser(UserIndex As Integer)
    Dim N As Integer

    N = RandomNumber(1, 10)

    If N < 3 Then
        UserList(UserIndex).flags.Envenenado = 1
        UserList(UserIndex).flags.EstasEnvenenado = Timer
        UserList(UserIndex).Counters.Veneno = Timer
        Call SendData(ToIndex, UserIndex, 0, "1P")
    End If

End Sub
Function SpawnNpc(NpcIndex As Integer, POS As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
    On Error GoTo Error
    Dim newpos As WorldPos
    Dim nIndex As Integer
    Dim PosicionValida As Boolean
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim it As Integer

    nIndex = OpenNPC(NpcIndex, Respawn)

    If nIndex > MAXNPCS Then
        SpawnNpc = nIndex
        Exit Function
    End If

    Do While Not PosicionValida
        Call ClosestLegalPos(POS, newpos)

        If LegalPos(newpos.Map, newpos.X, newpos.Y) Then
            Npclist(nIndex).POS.Map = newpos.Map
            Npclist(nIndex).POS.X = newpos.X
            Npclist(nIndex).POS.Y = newpos.Y
            PosicionValida = True
        Else
            newpos.X = 0
            newpos.Y = 0
        End If

        it = it + 1

        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(nIndex)
            SpawnNpc = MAXNPCS
            Call LogError("Más de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & POS.Map & " Index:" & NpcIndex)
            Exit Function
        End If
    Loop

    Map = newpos.Map
    X = Npclist(nIndex).POS.X
    Y = Npclist(nIndex).POS.Y

    Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)

    If NPCListable(nIndex) Then Call AgregarNPC(Npclist(nIndex).Numero, POS.Map)

    If FX Then
        Call SendData(ToNPCArea, nIndex, Npclist(NpcIndex).POS.Map, "TW" & SND_WARP)
        Call SendData(ToNPCArea, nIndex, Npclist(NpcIndex).POS.Map, "CFX" & Npclist(nIndex).Char.CharIndex & "," & FXWARP & "," & 0)
    End If

    SpawnNpc = nIndex

    Exit Function
Error:
    Call LogError("Error en SpawnNPC: " & Err.Description & " " & nIndex & " " & X & " " & Y)
End Function
Sub ReSpawnNpc(MiNPC As Npc)

    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.POS.Map, MiNPC.Orig)

End Sub
Function NPCHostiles(Map As Integer) As Integer
    Dim i As Integer
    Dim cont As Integer

    cont = 0

    For i = 1 To UBound(MapInfo(Map).NPCsTeoricos)
        cont = cont + MapInfo(Map).NPCsReales(i).Cantidad
    Next

    NPCHostiles = cont

End Function
Sub NPCTirarOro(MiNPC As Npc, UserIndex As Integer)
    Dim i As Integer, MiembroIndex As Integer

    If MiNPC.GiveGLD Then
        If UserList(UserIndex).PartyIndex = 0 Then
            If MiNPC.GiveGLD + UserList(UserIndex).Stats.GLD <= MAXORO Then
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiNPC.GiveGLD
                Call SendUserORO(UserIndex)
            End If
        Else
            For i = 1 To Party(UserList(UserIndex).PartyIndex).NroMiembros
                MiembroIndex = Party(UserList(UserIndex).PartyIndex).MiembrosIndex(i)
                If MiNPC.GiveGLD + UserList(MiembroIndex).Stats.GLD <= MAXORO Then
                    UserList(MiembroIndex).Stats.GLD = UserList(MiembroIndex).Stats.GLD + MiNPC.GiveGLD / Party(UserList(MiembroIndex).PartyIndex).NroMiembros
                    Call SendUserORO(MiembroIndex)
                End If
            Next
        End If
    End If

End Sub
Function NameNpc(Number As Integer) As String
    Dim A As Long, S As Long

    If Number > 499 Then
        A = Anpc_host
    Else
        A = ANpc
    End If

    S = INIBuscarSeccion(A, "NPC" & Number)

    NameNpc = INIDarClaveStr(A, S, "Name")

End Function
Function OpenNPC(NPCNumber As Integer, Optional ByVal Respawn = True) As Integer

    Dim NpcIndex As Integer

    Dim A As Long, S As Long

    If NPCNumber > 499 Then

        A = Anpc_host
    Else

        A = ANpc
    End If

    S = INIBuscarSeccion(A, "NPC" & NPCNumber)

    NpcIndex = NextOpenNPC

    If NpcIndex > MAXNPCS Then
        OpenNPC = NpcIndex
        Exit Function
    End If

    Npclist(NpcIndex).Numero = NPCNumber






    If S >= 0 Then
        Npclist(NpcIndex).Name = INIDarClaveStr(A, S, "Name")
        Npclist(NpcIndex).Desc = INIDarClaveStr(A, S, "Desc")

        Npclist(NpcIndex).Movement = INIDarClaveInt(A, S, "Movement")
        Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

        Npclist(NpcIndex).flags.AguaValida = INIDarClaveInt(A, S, "AguaValida")
        Npclist(NpcIndex).flags.TierraInvalida = INIDarClaveInt(A, S, "TierraInValida")
        Npclist(NpcIndex).flags.Faccion = INIDarClaveInt(A, S, "Faccion")

        Npclist(NpcIndex).NPCtype = INIDarClaveInt(A, S, "NpcType")

        Npclist(NpcIndex).Char.Body = INIDarClaveInt(A, S, "Body")
        Npclist(NpcIndex).Char.Head = INIDarClaveInt(A, S, "Head")
        Npclist(NpcIndex).Char.Heading = INIDarClaveInt(A, S, "Heading")

        Npclist(NpcIndex).Attackable = INIDarClaveInt(A, S, "Attackable")
        Npclist(NpcIndex).Comercia = INIDarClaveInt(A, S, "Comercia")
        Npclist(NpcIndex).Hostile = INIDarClaveInt(A, S, "Hostile")
        Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

        Npclist(NpcIndex).GiveEXP = INIDarClaveInt(A, S, "GiveEXP") * 100

        Npclist(NpcIndex).Veneno = INIDarClaveInt(A, S, "Veneno")

        Npclist(NpcIndex).flags.Domable = INIDarClaveInt(A, S, "Domable")

        Npclist(NpcIndex).MaxRecom = INIDarClaveInt(A, S, "MaxRecom")
        Npclist(NpcIndex).MinRecom = INIDarClaveInt(A, S, "MinRecom")
        Npclist(NpcIndex).Probabilidad = INIDarClaveInt(A, S, "Probabilidad")

        Npclist(NpcIndex).GiveGLD = INIDarClaveInt(A, S, "GiveGLD") * 100

        Npclist(NpcIndex).PoderAtaque = INIDarClaveInt(A, S, "PoderAtaque")
        Npclist(NpcIndex).PoderEvasion = INIDarClaveInt(A, S, "PoderEvasion")

        Npclist(NpcIndex).AutoCurar = INIDarClaveInt(A, S, "Autocurar")
        Npclist(NpcIndex).Stats.MaxHP = INIDarClaveInt(A, S, "MaxHP")
        Npclist(NpcIndex).Stats.MinHP = INIDarClaveInt(A, S, "MinHP")
        Npclist(NpcIndex).Stats.MaxHit = INIDarClaveInt(A, S, "MaxHIT")
        Npclist(NpcIndex).Stats.MinHit = INIDarClaveInt(A, S, "MinHIT")
        Npclist(NpcIndex).Stats.Def = INIDarClaveInt(A, S, "DEF")
        Npclist(NpcIndex).Stats.Alineacion = INIDarClaveInt(A, S, "Alineacion")
        Npclist(NpcIndex).Stats.ImpactRate = INIDarClaveInt(A, S, "ImpactRate")
        Npclist(NpcIndex).InvReSpawn = INIDarClaveInt(A, S, "InvReSpawn")


        Dim LoopC As Integer
        Dim ln As String
        Npclist(NpcIndex).Invent.NroItems = INIDarClaveInt(A, S, "NROITEMS")

        For LoopC = 1 To Minimo(30, Npclist(NpcIndex).Invent.NroItems)
            ln = INIDarClaveStr(A, S, "Obj" & LoopC)
            Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
        Next

        If Npclist(NpcIndex).InvReSpawn Or Npclist(NpcIndex).Comercia = 0 Then
            For LoopC = 1 To Minimo(30, Npclist(NpcIndex).Invent.NroItems)
                ln = INIDarClaveStr(A, S, "Obj" & LoopC)
                Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            Next
        End If

        Npclist(NpcIndex).flags.LanzaSpells = INIDarClaveInt(A, S, "LanzaSpells")
        If Npclist(NpcIndex).flags.LanzaSpells Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
        For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
            Npclist(NpcIndex).Spells(LoopC) = INIDarClaveInt(A, S, "Sp" & LoopC)
        Next


        If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
            Npclist(NpcIndex).NroCriaturas = INIDarClaveInt(A, S, "NroCriaturas")
            ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
            For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
                Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = INIDarClaveInt(A, S, "CI" & LoopC)
                Npclist(NpcIndex).Criaturas(LoopC).NpcName = INIDarClaveStr(A, S, "CN" & LoopC)

            Next
        End If


        Npclist(NpcIndex).Inflacion = INIDarClaveInt(A, S, "Inflacion")

        Npclist(NpcIndex).flags.NPCActive = True
        Npclist(NpcIndex).flags.UseAINow = False

        If Respawn Then
            Npclist(NpcIndex).flags.Respawn = INIDarClaveInt(A, S, "ReSpawn")
        Else
            Npclist(NpcIndex).flags.Respawn = 1
        End If

        Npclist(NpcIndex).flags.RespawnOrigPos = INIDarClaveInt(A, S, "OrigPos")
        Npclist(NpcIndex).flags.AfectaParalisis = INIDarClaveInt(A, S, "AfectaParalisis")
        Npclist(NpcIndex).flags.GolpeExacto = INIDarClaveInt(A, S, "GolpeExacto")
        Npclist(NpcIndex).flags.Apostador = INIDarClaveInt(A, S, "Apostador")
        Npclist(NpcIndex).flags.PocaParalisis = INIDarClaveInt(A, S, "PocaParalisis")
        Npclist(NpcIndex).flags.NoMagia = INIDarClaveInt(A, S, "NoMagia")
        Npclist(NpcIndex).VeInvis = INIDarClaveInt(A, S, "VerInvis")

        Npclist(NpcIndex).flags.Snd1 = INIDarClaveInt(A, S, "Snd1")
        Npclist(NpcIndex).flags.Snd2 = INIDarClaveInt(A, S, "Snd2")
        Npclist(NpcIndex).flags.Snd3 = INIDarClaveInt(A, S, "Snd3")
        Npclist(NpcIndex).flags.Snd4 = INIDarClaveInt(A, S, "Snd4")



        Dim aux As Long
        aux = INIDarClaveInt(A, S, "NROEXP")
        Npclist(NpcIndex).NroExpresiones = (aux)

        If aux Then
            ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
            For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
                Npclist(NpcIndex).Expresiones(LoopC) = INIDarClaveStr(A, S, "Exp" & LoopC)
            Next
        End If




        Npclist(NpcIndex).TipoItems = INIDarClaveInt(A, S, "TipoItems")
    End If


    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1



    OpenNPC = NpcIndex

End Function


Function OpenNPC_Viejo(NPCNumber As Integer, Optional ByVal Respawn = True) As Integer

    Dim NpcIndex As Integer
    Dim npcfile As String

    If NPCNumber > 499 Then
        npcfile = DatPath & "NPCs-HOSTILES.dat"
    Else
        npcfile = DatPath & "NPCs.dat"
    End If


    NpcIndex = NextOpenNPC

    If NpcIndex > MAXNPCS Then
        OpenNPC_Viejo = NpcIndex
        Exit Function
    End If

    Npclist(NpcIndex).Numero = NPCNumber
    Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NPCNumber, "Name")
    Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NPCNumber, "Desc")

    Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NPCNumber, "Movement"))
    Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

    Npclist(NpcIndex).flags.AguaValida = val(GetVar(npcfile, "NPC" & NPCNumber, "AguaValida"))
    Npclist(NpcIndex).flags.TierraInvalida = val(GetVar(npcfile, "NPC" & NPCNumber, "TierraInValida"))
    Npclist(NpcIndex).flags.Faccion = val(GetVar(npcfile, "NPC" & NPCNumber, "Faccion"))

    Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NPCNumber, "NpcType"))

    Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NPCNumber, "Body"))
    Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NPCNumber, "Head"))
    Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NPCNumber, "Heading"))

    Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NPCNumber, "Attackable"))
    Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NPCNumber, "Comercia"))
    Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NPCNumber, "Hostile"))
    Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile


    Npclist(NpcIndex).MaxRecom = val(GetVar(npcfile, "NPC" & NPCNumber, "MaxRecom"))
    Npclist(NpcIndex).MinRecom = val(GetVar(npcfile, "NPC" & NPCNumber, "MinRecom"))
    Npclist(NpcIndex).Probabilidad = val(GetVar(npcfile, "NPC" & NPCNumber, "Probabilidad"))


    Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NPCNumber, "GiveEXP"))

    Npclist(NpcIndex).Veneno = val(GetVar(npcfile, "NPC" & NPCNumber, "Veneno"))

    Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NPCNumber, "Domable"))


    Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NPCNumber, "GiveGLD"))

    Npclist(NpcIndex).PoderAtaque = val(GetVar(npcfile, "NPC" & NPCNumber, "PoderAtaque"))
    Npclist(NpcIndex).PoderEvasion = val(GetVar(npcfile, "NPC" & NPCNumber, "PoderEvasion"))

    Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NPCNumber, "InvReSpawn"))
    Npclist(NpcIndex).AutoCurar = val(GetVar(npcfile, "NPC" & NPCNumber, "autocurar"))


    Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NPCNumber, "MaxHP"))
    Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NPCNumber, "MinHP"))
    Npclist(NpcIndex).Stats.MaxHit = val(GetVar(npcfile, "NPC" & NPCNumber, "MaxHIT"))
    Npclist(NpcIndex).Stats.MinHit = val(GetVar(npcfile, "NPC" & NPCNumber, "MinHIT"))
    Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NPCNumber, "DEF"))
    Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NPCNumber, "Alineacion"))
    Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NPCNumber, "ImpactRate"))


    Dim LoopC As Integer
    Dim ln As String
    Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NPCNumber, "NROITEMS"))
    For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
        ln = GetVar(npcfile, "NPC" & NPCNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))

    Next

    Npclist(NpcIndex).flags.LanzaSpells = val(GetVar(npcfile, "NPC" & NPCNumber, "LanzaSpells"))
    If Npclist(NpcIndex).flags.LanzaSpells Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
    For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
        Npclist(NpcIndex).Spells(LoopC) = val(GetVar(npcfile, "NPC" & NPCNumber, "Sp" & LoopC))
    Next


    If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
        Npclist(NpcIndex).NroCriaturas = val(GetVar(npcfile, "NPC" & NPCNumber, "NroCriaturas"))
        ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
        For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
            Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = GetVar(npcfile, "NPC" & NPCNumber, "CI" & LoopC)
            Npclist(NpcIndex).Criaturas(LoopC).NpcName = GetVar(npcfile, "NPC" & NPCNumber, "CN" & LoopC)
        Next
    End If


    Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NPCNumber, "Inflacion"))

    Npclist(NpcIndex).flags.NPCActive = True
    Npclist(NpcIndex).flags.UseAINow = False

    If Respawn Then
        Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NPCNumber, "ReSpawn"))
    Else
        Npclist(NpcIndex).flags.Respawn = 1
    End If

    Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NPCNumber, "OrigPos"))
    Npclist(NpcIndex).flags.AfectaParalisis = val(GetVar(npcfile, "NPC" & NPCNumber, "AfectaParalisis"))
    Npclist(NpcIndex).flags.GolpeExacto = val(GetVar(npcfile, "NPC" & NPCNumber, "GolpeExacto"))
    Npclist(NpcIndex).flags.PocaParalisis = val(GetVar(npcfile, "NPC" & NPCNumber, "PocaParalisis"))
    Npclist(NpcIndex).VeInvis = val(GetVar(npcfile, "NPC" & NPCNumber, "veinvis"))



    Npclist(NpcIndex).flags.Snd1 = val(GetVar(npcfile, "NPC" & NPCNumber, "Snd1"))
    Npclist(NpcIndex).flags.Snd2 = val(GetVar(npcfile, "NPC" & NPCNumber, "Snd2"))
    Npclist(NpcIndex).flags.Snd3 = val(GetVar(npcfile, "NPC" & NPCNumber, "Snd3"))
    Npclist(NpcIndex).flags.Snd4 = val(GetVar(npcfile, "NPC" & NPCNumber, "Snd4"))



    Dim aux As String
    aux = GetVar(npcfile, "NPC" & NPCNumber, "NROEXP")
    If Len(aux) = 0 Then
        Npclist(NpcIndex).NroExpresiones = 0
    Else
        Npclist(NpcIndex).NroExpresiones = val(aux)
        ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
        For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
            Npclist(NpcIndex).Expresiones(LoopC) = GetVar(npcfile, "NPC" & NPCNumber, "Exp" & LoopC)
        Next
    End If




    Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NPCNumber, "TipoItems"))


    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1



    OpenNPC_Viejo = NpcIndex

End Function

Sub EnviarListaCriaturas(UserIndex As Integer, NpcIndex)
    Dim SD As String
    Dim k As Integer
    SD = SD & Npclist(NpcIndex).NroCriaturas & ","
    For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
    Next
    SD = "LSTCRI" & SD
    Call SendData(ToIndex, UserIndex, 0, SD)
End Sub


Sub DoFollow(NpcIndex As Integer, UserIndex As Integer)

    If Npclist(NpcIndex).flags.Follow Then
        Npclist(NpcIndex).flags.AttackedBy = 0
        Npclist(NpcIndex).flags.Follow = False
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    Else
        Npclist(NpcIndex).flags.AttackedBy = UserIndex
        Npclist(NpcIndex).flags.Follow = True
        Npclist(NpcIndex).Movement = 4
        Npclist(NpcIndex).Hostile = 0
    End If

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).flags.Follow = True
    Npclist(NpcIndex).Movement = SIGUE_AMO
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNpc = 0

End Sub

