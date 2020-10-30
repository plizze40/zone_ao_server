Attribute VB_Name = "UsUaRiOs"
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
Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

    If UserList(AttackerIndex).POS.Map <> 8 Then
        Dim DaExp As Integer
        DaExp = CInt(UserList(VictimIndex).Stats.ELV * RandomNumber(1, 4))
        Call AddtoVar(UserList(AttackerIndex).Stats.Exp, DaExp, MAXEXP)
    End If

    Call SendData(ToIndex, AttackerIndex, 0, "1Q" & UserList(VictimIndex).Name)
    Call SendData(ToIndex, AttackerIndex, 0, "EX" & DaExp)
    Call SendData(ToIndex, VictimIndex, 0, "1R" & UserList(AttackerIndex).Name)

    Call UserDie(VictimIndex)

End Sub
Sub RevivirUsuarioNPC(UserIndex As Integer)

    UserList(UserIndex).flags.Muerto = 0
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

    Call DarCuerpoDesnudo(UserIndex)
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendUserStatsBox(UserIndex)

End Sub
Sub RevivirUsuario(ByVal Resucitador As Integer, UserIndex As Integer, ByVal Lleno As Boolean)

    UserList(Resucitador).Stats.MinSta = 0
    UserList(Resucitador).Stats.MinAGU = 0
    UserList(Resucitador).Stats.MinHam = 0
    UserList(Resucitador).flags.Sed = 1
    UserList(Resucitador).flags.Hambre = 1

    UserList(UserIndex).flags.Muerto = 0

    If Lleno Then
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        UserList(UserIndex).flags.Hambre = 0
    Else
        UserList(UserIndex).Stats.MinHP = 1
        UserList(UserIndex).Stats.MinSta = 0
        UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MinHam = 0
        UserList(UserIndex).Stats.MinAGU = 0
        UserList(UserIndex).flags.Sed = 1
        UserList(UserIndex).flags.Hambre = 1
    End If

    Call DarCuerpoDesnudo(UserIndex)
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

    Call SendUserStatsBox(Resucitador)
    Call EnviarHambreYsed(Resucitador)

    Call SendUserStatsBox(UserIndex)
    Call EnviarHambreYsed(UserIndex)

End Sub
Sub ReNombrar(UserIndex As Integer, NewNick As String)

    If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) Then
        Kill CharPath & UCase$(UserList(UserIndex).Name) & ".chr"
    End If

    Call SendData(ToAll, 0, 0, "||El usuario " & UserList(UserIndex).Name & " ha sido rebautizado como " & NewNick & "." & FONTTYPE_FIGHT)
    UserList(UserIndex).Name = NewNick
    Call WarpUserChar(UserIndex, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y, False)

End Sub
Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, _
                   ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                   ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

    On Error Resume Next

    UserList(UserIndex).Char.Body = Body
    UserList(UserIndex).Char.Head = Head
    UserList(UserIndex).Char.Heading = Heading
    UserList(UserIndex).Char.WeaponAnim = Arma
    UserList(UserIndex).Char.ShieldAnim = Escudo
    UserList(UserIndex).Char.CascoAnim = Casco

    Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco)

End Sub
Sub ChangeUserCharB(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, _
                    ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

    On Error Resume Next

    UserList(UserIndex).Char.Body = Body
    UserList(UserIndex).Char.Head = Head
    UserList(UserIndex).Char.Heading = Heading
    UserList(UserIndex).Char.WeaponAnim = Arma
    UserList(UserIndex).Char.ShieldAnim = Escudo
    UserList(UserIndex).Char.CascoAnim = Casco

    Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco & "," & UserList(UserIndex).flags.Navegando)

End Sub
Sub ChangeUserCasco(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, _
                    ByVal Casco As Integer)

    On Error Resume Next

    If UserList(UserIndex).Char.CascoAnim <> Casco Then
        UserList(UserIndex).Char.CascoAnim = Casco
        Call SendData(sndRoute, sndIndex, sndMap, "7C" & UserList(UserIndex).Char.CharIndex & "," & Casco)
    End If

End Sub
Sub ChangeUserEscudo(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, ByVal Escudo As Integer)
    On Error Resume Next

    If UserList(UserIndex).Char.ShieldAnim <> Escudo Then
        UserList(UserIndex).Char.ShieldAnim = Escudo
        Call SendData(sndRoute, sndIndex, sndMap, "6C" & UserList(UserIndex).Char.CharIndex & "," & Escudo)
    End If

End Sub


Sub ChangeUserArma(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, _
                   ByVal Arma As Integer)

    On Error Resume Next

    If UserList(UserIndex).Char.WeaponAnim <> Arma Then
        UserList(UserIndex).Char.WeaponAnim = Arma
        Call SendData(sndRoute, sndIndex, sndMap, "5C" & UserList(UserIndex).Char.CharIndex & "," & Arma)
    End If


End Sub


Sub ChangeUserHead(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, _
                   ByVal Head As Integer)

    On Error Resume Next

    If UserList(UserIndex).Char.Head <> Head Then
        UserList(UserIndex).Char.Head = Head
        Call SendData(sndRoute, sndIndex, sndMap, "4C" & UserList(UserIndex).Char.CharIndex & "," & Head)
    End If

End Sub

Sub ChangeUserBody(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, _
                   ByVal Body As Integer)

    On Error Resume Next
    UserList(UserIndex).Char.Body = Body
    Call SendData(sndRoute, sndIndex, sndMap, "3C" & UserList(UserIndex).Char.CharIndex & "," & Body)


End Sub
Sub ChangeUserHeading(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, _
                      ByVal Heading As Byte)
    On Error Resume Next

    UserList(UserIndex).Char.Heading = Heading
    Call SendData(sndRoute, sndIndex, sndMap, "2C" & UserList(UserIndex).Char.CharIndex & "," & Heading)

End Sub
Sub EnviarSubirNivel(UserIndex As Integer, ByVal Puntos As Integer)

    Call SendData(ToIndex, UserIndex, 0, "SUNI" & Puntos)

End Sub
Sub EnviarSkills(UserIndex As Integer)
    Dim i As Integer
    Dim cad As String

    For i = 1 To NUMSKILLS
        cad = cad & UserList(UserIndex).Stats.UserSkills(i) & ","
    Next

    SendData ToIndex, UserIndex, 0, "SKILLS" & cad

End Sub
Sub EnviarFama(UserIndex As Integer)
    Dim cad As String

    cad = UserList(UserIndex).Faccion.Quests & ","
    cad = cad & UserList(UserIndex).Faccion.torneos & ","

    If EsNewbie(UserIndex) Then
        cad = cad & UserList(UserIndex).Faccion.Matados(Caos) & ","
        cad = cad & UserList(UserIndex).Faccion.Matados(Neutral)

        Call SendData(ToIndex, UserIndex, 0, "FAMA3," & cad)
    Else
        Select Case UserList(UserIndex).Faccion.Bando
        Case Neutral
            cad = cad & UserList(UserIndex).Faccion.BandoOriginal & ","
            cad = cad & UserList(UserIndex).Faccion.Matados(Real) & ","
            cad = cad & UserList(UserIndex).Faccion.Matados(Caos) & ","

        Case Real, Caos
            cad = cad & Titulo(UserIndex) & ","
            cad = cad & UserList(UserIndex).Faccion.Matados(Enemigo(UserList(UserIndex).Faccion.Bando)) & ","

        End Select
        cad = cad & UserList(UserIndex).Faccion.Matados(Neutral)
        Call SendData(ToIndex, UserIndex, 0, "FAMA" & UserList(UserIndex).Faccion.Bando & "," & cad)
    End If

End Sub
Function GeneroLetras(Genero As Byte) As String

    If Genero = 1 Then
        GeneroLetras = "Mujer"
    Else
        GeneroLetras = "Hombre"
    End If

End Function
Sub EnviarMiniSt(UserIndex As Integer)
    Dim cad As String

    cad = cad & UserList(UserIndex).Stats.VecesMurioUsuario & ","
    cad = cad & UserList(UserIndex).Faccion.Matados(Caos) & ","
    cad = cad & UserList(UserIndex).Stats.NPCsMuertos & ","
    cad = cad & UserList(UserIndex).Faccion.Matados(Neutral) + UserList(UserIndex).Faccion.Matados(Real) + UserList(UserIndex).Faccion.Matados(Caos) & ","
    cad = cad & ListaClases(UserList(UserIndex).Clase) & ","
    cad = cad & ListaRazas(UserList(UserIndex).Raza) & ","
    cad = cad & UserList(UserIndex).Faccion.Matados(Real) & ","

    Call SendData(ToIndex, UserIndex, 0, "MIST" & cad)

End Sub
Sub EnviarAtrib(UserIndex As Integer)
    Dim i As Integer
    Dim cad As String

    For i = 1 To NUMATRIBUTOS
        cad = cad & UserList(UserIndex).Stats.UserAtributos(i) & ","
    Next

    Call SendData(ToIndex, UserIndex, 0, "ATR" & cad)

End Sub
Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)

    On Error GoTo ErrorHandler

    CharList(UserList(UserIndex).Char.CharIndex) = 0

    If UserList(UserIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If

    MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).UserIndex = 0


    Call SendData(ToMap, UserIndex, UserList(UserIndex).POS.Map, "BP" & UserList(UserIndex).Char.CharIndex)

    UserList(UserIndex).Char.CharIndex = 0

    NumChars = NumChars - 1

    Exit Sub

ErrorHandler:
    Call LogError("Error en EraseUserchar")

End Sub
Sub UpdateUserChar(UserIndex As Integer)
    On Error Resume Next
    Dim bCr As Byte
    Dim info As String

    If UserList(UserIndex).flags.Privilegios Then
        bCr = 1
    ElseIf UserList(UserIndex).POS.Map = 88 Then
        bCr = 11
    ElseIf UserList(UserIndex).Faccion.Bando = Real Then
        bCr = 2
    ElseIf UserList(UserIndex).Faccion.Bando = Caos Then
        bCr = 3
    ElseIf EsNewbie(UserIndex) Then
        bCr = 4
    Else: bCr = 5
    End If

    info = "PW" & UserList(UserIndex).Char.CharIndex & "," & bCr & "," & UserList(UserIndex).Name

    If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then info = info & " <" & UserList(UserIndex).GuildInfo.GuildName & ">"

    Call SendData(ToMap, UserIndex, UserList(UserIndex).POS.Map, (info))

End Sub
Sub MakeUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer, Map As Integer, X As Integer, Y As Integer)
    On Error Resume Next
    Dim CharIndex As Integer

    If Not InMapBounds(X, Y) Then Exit Sub


    If UserList(UserIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        UserList(UserIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = UserIndex
    End If


    MapData(Map, X, Y).UserIndex = UserIndex


    Dim klan$
    klan$ = UserList(UserIndex).GuildInfo.GuildName
    Dim bCr As Byte
    If UserList(UserIndex).flags.Privilegios Then
        bCr = 1
    ElseIf UserList(UserIndex).POS.Map = 88 Then
        bCr = 11
    ElseIf UserList(UserIndex).Faccion.Bando = Real And UserList(UserIndex).flags.EsConseReal = 0 Then
        bCr = 2
    ElseIf UserList(UserIndex).Faccion.Bando = Caos And UserList(UserIndex).flags.EsConseCaos = 0 Then
        bCr = 3
    ElseIf EsNewbie(UserIndex) Then
        bCr = 4
    ElseIf UserList(UserIndex).flags.EsConseCaos And UserList(UserIndex).Faccion.Bando = Caos Then
        bCr = 6
    ElseIf UserList(UserIndex).flags.EsConseReal And UserList(UserIndex).Faccion.Bando = Real Then
        bCr = 7
    ElseIf UserList(UserIndex).flags.EsConcilioNegro And UserList(UserIndex).Faccion.Bando = Neutral Then
        bCr = 8
    Else
        bCr = 5
    End If

    If Len(klan$) > 0 Then klan = " <" & klan$ & ">"

    Call SendData(sndRoute, sndIndex, sndMap, ("CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & klan$ & "," & bCr & "," & UserList(UserIndex).flags.Invisible))

    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).Char.loops = LoopAdEternum
        If UserList(UserIndex).Stats.ELV < 15 Then
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
            UserList(UserIndex).Char.FX = FXMEDITARCHICO
        ElseIf UserList(UserIndex).Stats.ELV < 30 Then
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
            UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
        ElseIf UserList(UserIndex).Stats.ELV < 45 Then
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
            UserList(UserIndex).Char.FX = FXMEDITARGRANDE
        ElseIf UserList(UserIndex).Stats.ELV < 46 Then
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGIGANTE & "," & LoopAdEternum)
            UserList(UserIndex).Char.FX = FXMEDITARGIGANTE
        End If
    End If

End Sub
Function Redondea(ByVal Number As Single) As Integer

    If Number > Fix(Number) Then
        Redondea = Fix(Number) + 1
    Else: Redondea = Number
    End If

End Function
Sub CheckUserLevel(UserIndex As Integer)
    On Error GoTo errhandler
    Dim Pts As Integer
    Dim SubeHit As Integer
    Dim AumentoST As Integer
    Dim AumentoMANA As Integer
    Dim WasNewbie As Boolean

    Do Until UserList(UserIndex).Stats.Exp < UserList(UserIndex).Stats.ELU
        If UserList(UserIndex).Stats.ELV >= STAT_MAXELV Then
            UserList(UserIndex).Stats.Exp = 0
            UserList(UserIndex).Stats.ELU = 0
            Call SendData(ToIndex, UserIndex, 0, "ELV")
            Exit Sub
        End If

        WasNewbie = EsNewbie(UserIndex)

        If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then

            If UserList(UserIndex).Stats.ELV >= 14 And ClaseBase(UserList(UserIndex).Clase) Then
                Call SendData(ToIndex, UserIndex, 0, "!6")
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.ELU - 1
                Call SendUserEXP(UserIndex)
                Exit Sub
            End If

            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SOUND_NIVEL)

            Call SendData(ToIndex, UserIndex, 0, "1S" & UserList(UserIndex).Stats.ELV + 1)

            If UserList(UserIndex).Stats.ELV = 1 Then
                Pts = 15
            Else
                Pts = 15
            End If

            UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts

            Call SendData(ToIndex, UserIndex, 0, "1T" & Pts)

            UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp - UserList(UserIndex).Stats.ELU
            UserList(UserIndex).Stats.ELU = ELUs(UserList(UserIndex).Stats.ELV)

            Dim AumentoHP As Integer
            Dim SubePromedio As Single

            SubePromedio = UserList(UserIndex).Stats.UserAtributos(Constitucion) / 2 - Resta(UserList(UserIndex).Clase)
            AumentoHP = RandomNumber(Fix(SubePromedio - 1), Redondea(SubePromedio + 1))
            SubeHit = AumentoHit(UserList(UserIndex).Clase)

            Select Case UserList(UserIndex).Clase
            Case CIUDADANO, TRABAJADOR, EXPERTO_MINERALES
                AumentoST = 15

            Case MINERO
                AumentoST = 15 + AdicionalSTMinero

            Case HERRERO
                AumentoST = 15

            Case EXPERTO_MADERA
                AumentoST = 15

            Case TALADOR
                AumentoST = 15 + AdicionalSTLeñador

            Case CARPINTERO
                AumentoST = 15

            Case PESCADOR
                AumentoST = 15 + AdicionalSTPescador

            Case SASTRE
                AumentoST = 15

            Case HECHICERO
                AumentoST = 15
                AumentoMANA = 2.2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            Case MAGO
                AumentoST = Maximo(5, 15 - AdicionalSTLadron / 2)
                Select Case UserList(UserIndex).Stats.MaxMAN
                Case Is < 2300
                    AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                Case Is < 2500
                    AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                Case Else
                    AumentoMANA = 1.5 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                End Select

                If UserList(UserIndex).Stats.ELV > 45 Then AumentoMANA = 0

            Case NIGROMANTE
                AumentoST = Maximo(5, 15 - AdicionalSTLadron / 2)
                AumentoMANA = 2.2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            Case ORDEN_SAGRADA
                AumentoST = 15
                AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            Case PALADIN
                AumentoST = 15
                AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)

                If UserList(UserIndex).Stats.MaxHit >= 99 Then SubeHit = 1

            Case CLERIGO
                AumentoST = 15
                AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            Case NATURALISTA
                AumentoST = 15
                AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            Case BARDO
                AumentoST = 15
                AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            Case DRUIDA
                AumentoST = 15
                AumentoMANA = 2.2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            Case SIGILOSO
                AumentoST = 15
                AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)

            Case ASESINO
                AumentoST = 15
                AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)

                If UserList(UserIndex).Stats.MaxHit >= 99 Then SubeHit = 1

            Case CAZADOR
                AumentoST = 15
                AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)

                If UserList(UserIndex).Stats.MaxHit >= 99 Then SubeHit = 1

            Case SIN_MANA
                AumentoST = 15

            Case CABALLERO
                AumentoST = 15

            Case ARQUERO
                AumentoST = 15

                If UserList(UserIndex).Stats.MaxHit >= 99 Then SubeHit = 2

            Case GUERRERO
                AumentoST = 15

                If UserList(UserIndex).Stats.MaxHit >= 99 Then SubeHit = 2

            Case BANDIDO
                AumentoST = 15

            Case PIRATA
                AumentoST = 15

            Case LADRON
                AumentoST = 15

            Case Else
                AumentoST = 15 + AdicionalSTLadron

            End Select

            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            UserList(UserIndex).Stats.MaxSta = UserList(UserIndex).Stats.MaxSta + AumentoST

            Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 2200 + 800 * Buleano(UserList(UserIndex).Clase And UserList(UserIndex).Recompensas(2) = 2))
            UserList(UserIndex).Stats.MaxHit = UserList(UserIndex).Stats.MaxHit + SubeHit
            UserList(UserIndex).Stats.MinHit = UserList(UserIndex).Stats.MinHit + SubeHit

            Call SendData(ToIndex, UserIndex, 0, "1U" & AumentoHP & "," & AumentoST & "," & AumentoMANA & "," & SubeHit)

            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

            Call EnviarSkills(UserIndex)
            Call EnviarSubirNivel(UserIndex, Pts)

            Call SendUserStatsBox(UserIndex)

            If Not EsNewbie(UserIndex) And WasNewbie Then
                If UserList(UserIndex).POS.Map = 1 Or UserList(UserIndex).POS.Map = 49 Then
                    Call WarpUserChar(UserIndex, 1, 50, 50, True)
                Else
                    Call UpdateUserChar(UserIndex)
                End If
                Call QuitarNewbieObj(UserIndex)
                Call SendData(ToIndex, UserIndex, 0, "SUFA1")
            End If

            Call CheckUserLevel(UserIndex)

        Else

            Call SendUserEXP(UserIndex)

        End If


        If PuedeSubirClase(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SUCL1")
        If PuedeRecompensa(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SURE1")

    Loop

    Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub
Function PuedeRecompensa(UserIndex As Integer) As Byte

    If UserList(UserIndex).Clase = SASTRE Then Exit Function

    If UserList(UserIndex).Recompensas(1) = 0 And UserList(UserIndex).Stats.ELV >= 18 Then
        PuedeRecompensa = 1
        Exit Function
    End If

    If UserList(UserIndex).Clase = TALADOR Or UserList(UserIndex).Clase = PESCADOR Then Exit Function

    If UserList(UserIndex).Stats.ELV >= 25 And UserList(UserIndex).Recompensas(2) = 0 Then
        PuedeRecompensa = 2
        Exit Function
    End If

    If UserList(UserIndex).Clase = CARPINTERO Then Exit Function

    If UserList(UserIndex).Recompensas(3) = 0 And _
       (UserList(UserIndex).Stats.ELV >= 34 Or _
        (ClaseTrabajadora(UserList(UserIndex).Clase) And UserList(UserIndex).Stats.ELV >= 32) Or _
        ((UserList(UserIndex).Clase = PIRATA Or UserList(UserIndex).Clase = LADRON) And UserList(UserIndex).Stats.ELV >= 30)) Then
        PuedeRecompensa = 3
        Exit Function
    End If

End Function
Function PuedeFaccion(UserIndex As Integer) As Boolean

    PuedeFaccion = Not EsNewbie(UserIndex) And UserList(UserIndex).Faccion.BandoOriginal = Neutral And Len(UserList(UserIndex).GuildInfo.GuildName) = 0 And UserList(UserIndex).flags.Privilegios = 0

End Function
Function PuedeSubirClase(UserIndex As Integer) As Boolean

    PuedeSubirClase = (UserList(UserIndex).Stats.ELV >= 3 And UserList(UserIndex).Clase = CIUDADANO) Or _
                      (UserList(UserIndex).Stats.ELV >= 6 And (UserList(UserIndex).Clase = LUCHADOR Or UserList(UserIndex).Clase = TRABAJADOR)) Or _
                      (UserList(UserIndex).Stats.ELV >= 9 And (UserList(UserIndex).Clase = EXPERTO_MINERALES Or UserList(UserIndex).Clase = EXPERTO_MADERA Or UserList(UserIndex).Clase = CON_MANA Or UserList(UserIndex).Clase = SIN_MANA)) Or _
                      (UserList(UserIndex).Stats.ELV >= 12 And (UserList(UserIndex).Clase = CABALLERO Or UserList(UserIndex).Clase = BANDIDO Or UserList(UserIndex).Clase = HECHICERO Or UserList(UserIndex).Clase = NATURALISTA Or UserList(UserIndex).Clase = ORDEN_SAGRADA Or UserList(UserIndex).Clase = SIGILOSO))

End Function
Function PuedeAtravesarAgua(UserIndex As Integer) As Boolean

    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1

End Function
Private Sub EnviaNuevaPosUsuarioPj(UserIndex As Integer, ByVal Quien As Integer)

    Call SendData(ToIndex, UserIndex, 0, ("LP" & UserList(Quien).Char.CharIndex & "," & UserList(Quien).POS.X & "," & UserList(Quien).POS.Y & "," & UserList(Quien).Char.Heading))

End Sub
Private Sub EnviaNuevaPosNPC(UserIndex As Integer, NpcIndex As Integer)

    Call SendData(ToIndex, UserIndex, 0, ("LP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).POS.X & "," & Npclist(NpcIndex).POS.Y & "," & Npclist(NpcIndex).Char.Heading))

End Sub
Sub CalcularValores(UserIndex As Integer)
    Dim SubePromedio As Single
    Dim HPReal As Integer
    Dim HitReal As Integer
    Dim i As Integer

    HPReal = 15 + RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 3)
    HitReal = AumentoHit(UserList(UserIndex).Clase) * UserList(UserIndex).Stats.ELV
    SubePromedio = UserList(UserIndex).Stats.UserAtributos(Constitucion) / 2 - Resta(UserList(UserIndex).Clase)

    For i = 1 To UserList(UserIndex).Stats.ELV - 1
        HPReal = HPReal + RandomNumber(Redondea(SubePromedio - 2), Fix(SubePromedio + 2))
    Next

    Call CalcularMana(UserIndex)

    UserList(UserIndex).Stats.MinHit = HitReal
    UserList(UserIndex).Stats.MaxHit = HitReal + 1

    UserList(UserIndex).Stats.MinHP = Minimo(UserList(UserIndex).Stats.MinHP, HPReal)
    UserList(UserIndex).Stats.MaxHP = HPReal
    Call SendUserStatsBox(UserIndex)

End Sub
Sub CalcularMana(UserIndex As Integer)
    Dim ManaReal As Integer

    Select Case (UserList(UserIndex).Clase)
    Case HECHICERO
        ManaReal = 100 + 2.2 * (UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1))

    Case MAGO
        ManaReal = 100 + 3 * (UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1))

    Case ORDEN_SAGRADA
        ManaReal = UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1)

    Case CLERIGO
        ManaReal = 50 + 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1)

    Case NATURALISTA
        ManaReal = 50 + 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1)

    Case DRUIDA
        ManaReal = 50 + 2.1 * UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1)

    Case SIGILOSO
        ManaReal = 50 + UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1)
    End Select

    If ManaReal Then
        UserList(UserIndex).Stats.MinMAN = Minimo(UserList(UserIndex).Stats.MinMAN, ManaReal)
        UserList(UserIndex).Stats.MaxMAN = ManaReal
    End If

End Sub
Private Sub EnviaGenteEnNuevoRango(UserIndex As Integer, ByVal nHeading As Byte)
    Dim X As Integer, Y As Integer
    Dim M As Integer

    M = UserList(UserIndex).POS.Map

    Select Case nHeading

    Case NORTH, SOUTH

        If nHeading = NORTH Then
            Y = UserList(UserIndex).POS.Y - MinYBorder - 3
        Else
            Y = UserList(UserIndex).POS.Y + MinYBorder + 3
        End If
        For X = UserList(UserIndex).POS.X - MinXBorder - 2 To UserList(UserIndex).POS.X + MinXBorder + 2
            If MapData(M, X, Y).UserIndex Then
                Call EnviaNuevaPosUsuarioPj(UserIndex, MapData(M, X, Y).UserIndex)
            ElseIf MapData(M, X, Y).NpcIndex Then
                Call EnviaNuevaPosNPC(UserIndex, MapData(M, X, Y).NpcIndex)
            End If
        Next
    Case EAST, WEST

        If nHeading = EAST Then
            X = UserList(UserIndex).POS.X + MinXBorder + 3
        Else
            X = UserList(UserIndex).POS.X - MinXBorder - 3
        End If
        For Y = UserList(UserIndex).POS.Y - MinYBorder - 2 To UserList(UserIndex).POS.Y + MinYBorder + 2
            If MapData(M, X, Y).UserIndex Then
                Call EnviaNuevaPosUsuarioPj(UserIndex, MapData(M, X, Y).UserIndex)
            ElseIf MapData(M, X, Y).NpcIndex Then
                Call EnviaNuevaPosNPC(UserIndex, MapData(M, X, Y).NpcIndex)
            End If
        Next
    End Select

End Sub
Sub CancelarSacrificio(Sacrificado As Integer)
    Dim Sacrificador As Integer

    Sacrificador = UserList(Sacrificado).flags.Sacrificador

    UserList(Sacrificado).flags.Sacrificando = 0
    UserList(Sacrificado).flags.Sacrificador = 0
    UserList(Sacrificador).flags.Sacrificado = 0

    Call SendData(ToIndex, Sacrificado, 0, "||¡El sacrificio fue cancelado!" & FONTTYPE_INFO)
    Call SendData(ToIndex, Sacrificador, 0, "||¡El sacrificio fue cancelado!" & FONTTYPE_INFO)

End Sub
Sub MoveUserChar(UserIndex As Integer, ByVal nHeading As Byte)
    On Error Resume Next
    Dim nPos As WorldPos



    UserList(UserIndex).Counters.Pasos = UserList(UserIndex).Counters.Pasos + 1

    nPos = UserList(UserIndex).POS
    Call HeadtoPos(nHeading, nPos)

    If UserList(UserIndex).flags.Sacrificado > 0 Then Call CancelarSacrificio(UserList(UserIndex).flags.Sacrificado)
    If UserList(UserIndex).flags.Sacrificando = 1 Then Call CancelarSacrificio(UserIndex)

    If Not LegalPos(UserList(UserIndex).POS.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(UserIndex)) Then
        Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).POS.X & "," & UserList(UserIndex).POS.Y)
        If MapData(nPos.Map, nPos.X, nPos.Y).UserIndex Then
            Call EnviaNuevaPosUsuarioPj(UserIndex, MapData(nPos.Map, nPos.X, nPos.Y).UserIndex)
        ElseIf MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex Then
            Call EnviaNuevaPosNPC(UserIndex, MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex)
        End If
        Exit Sub
    End If

    Call SendData(ToPCAreaButIndexG, UserIndex, UserList(UserIndex).POS.Map, ("MP" & UserList(UserIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y))
    Call EnviaGenteEnNuevoRango(UserIndex, nHeading)
    MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).UserIndex = 0
    UserList(UserIndex).POS = nPos
    UserList(UserIndex).Char.Heading = nHeading
    MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).UserIndex = UserIndex
    Call DoTileEvents(UserIndex)

End Sub
Sub DesequiparItem(UserIndex As Integer, Slot As Byte)

    Call SendData(ToIndex, UserIndex, 0, "8J" & Slot)

End Sub
Sub EquiparItem(UserIndex As Integer, Slot As Byte)

    Call SendData(ToIndex, UserIndex, 0, "7J" & Slot)

End Sub

Sub SendUserItem(UserIndex As Integer, Slot As Byte, JustAmount As Boolean)
    Dim MiObj As UserOBJ
    Dim info As String

    MiObj = UserList(UserIndex).Invent.Object(Slot)

    If MiObj.OBJIndex Then
        If Not JustAmount Then
            info = "CSI" & Slot & "," & ObjData(MiObj.OBJIndex).Name & "," & MiObj.Amount & "," & MiObj.Equipped & "," & ObjData(MiObj.OBJIndex).GrhIndex & "," _
                 & ObjData(MiObj.OBJIndex).ObjType & "," & Round(ObjData(MiObj.OBJIndex).Valor / 3)
            Select Case ObjData(MiObj.OBJIndex).ObjType
            Case OBJTYPE_WEAPON
                info = info & "," & ObjData(MiObj.OBJIndex).MaxHit & "," & ObjData(MiObj.OBJIndex).MinHit
            Case OBJTYPE_ARMOUR
                info = info & "," & ObjData(MiObj.OBJIndex).SubTipo & "," & ObjData(MiObj.OBJIndex).MaxDef & "," & ObjData(MiObj.OBJIndex).MinDef
            Case OBJTYPE_POCIONES
                info = info & "," & ObjData(MiObj.OBJIndex).TipoPocion & "," & ObjData(MiObj.OBJIndex).MaxModificador & "," & ObjData(MiObj.OBJIndex).MinModificador
            End Select
            Call SendData(ToIndex, UserIndex, 0, info)
        Else: Call SendData(ToIndex, UserIndex, 0, "CSO" & Slot & "," & MiObj.Amount)
        End If
    Else: Call SendData(ToIndex, UserIndex, 0, "2H" & Slot)
    End If

End Sub
Function NextOpenCharIndex() As Integer
    Dim LoopC As Integer

    For LoopC = 1 To LastChar + 1
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            If LoopC > LastChar Then LastChar = LoopC
            Exit Function
        End If
    Next

End Function
Function NextOpenUser() As Integer
    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1) Then Exit For
    Next

    NextOpenUser = LoopC

End Function

Sub SendUserStatsBox(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.Exp & "," & UserList(UserIndex).POS.Map)
End Sub
Sub SendUserHP(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "5A" & UserList(UserIndex).Stats.MinHP)
End Sub
Sub SendUserMANA(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "5D" & UserList(UserIndex).Stats.MinMAN)
End Sub
Sub SendUserMAXHP(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "8B" & UserList(UserIndex).Stats.MaxHP)
End Sub
Sub SendUserMAXMANA(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "9B" & UserList(UserIndex).Stats.MaxMAN)
End Sub
Sub SendUserSTA(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "5E" & UserList(UserIndex).Stats.MinSta)
End Sub
Sub SendUserORO(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "5F" & UserList(UserIndex).Stats.GLD)
End Sub
Sub SendUserEXP(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "5G" & UserList(UserIndex).Stats.Exp)
End Sub
Sub SendUserMANASTA(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "5H" & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MinSta)
End Sub
Sub SendUserHPSTA(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "5I" & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MinSta)
End Sub
Sub EnviarHambreYsed(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "EHYS" & UserList(UserIndex).Stats.MaxAGU & "," & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MaxHam & "," & UserList(UserIndex).Stats.MinHam)
End Sub
Sub EnviarHyS(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "5J" & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MinHam)
End Sub

Sub SendUserSTAtsTxt(ByVal sendIndex As Integer, UserIndex As Integer)

    Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).Name & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(UserIndex).Stats.FIT & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "||Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & FONTTYPE_INFO)

    If UserList(UserIndex).Invent.WeaponEqpObjIndex Then
        Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHit & "/" & UserList(UserIndex).Stats.MaxHit & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHit & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHit & ")" & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHit & "/" & UserList(UserIndex).Stats.MaxHit & FONTTYPE_INFO)
    End If

    Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef + 2 * Buleano(UserList(UserIndex).Clase = GUERRERO And UserList(UserIndex).Recompensas(2) = 2) & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef + 2 * Buleano(UserList(UserIndex).Clase = GUERRERO And UserList(UserIndex).Recompensas(2) = 2) & FONTTYPE_INFO)

    If UserList(UserIndex).Invent.CascoEqpObjIndex Then
        Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If

    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        Call SendData(ToIndex, sendIndex, 0, "||(ESCUDO) Defensa extra: " & ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MinDef & " / " & ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MaxDef & FONTTYPE_INFO)
    End If

    If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then
        Call SendData(ToIndex, sendIndex, 0, "||Clan: " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, sendIndex, 0, "||Status: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    Call SendData(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
        If UserList(UserIndex).GuildInfo.ClanFundado = UserList(UserIndex).GuildInfo.GuildName Then
            Call SendData(ToIndex, sendIndex, 0, "||Status: " & "Fundador/Lider" & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, sendIndex, 0, "||Status: " & "Lider" & FONTTYPE_INFO)
        End If
    Else
        Call SendData(ToIndex, sendIndex, 0, "||Status: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    Call SendData(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)

    Call SendData(ToIndex, sendIndex, 0, "||Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).POS.X & "," & UserList(UserIndex).POS.Y & " en mapa " & UserList(UserIndex).POS.Map & FONTTYPE_INFO)

    Call SendData(ToIndex, sendIndex, 0, "||Ciudadanos matados: " & UserList(UserIndex).Faccion.Matados(Real) & " / Criminales matados: " & UserList(UserIndex).Faccion.Matados(Caos) & " / Neutrales matados: " & UserList(UserIndex).Faccion.Matados(Neutral) & FONTTYPE_INFO)

End Sub
Sub SendUserInvTxt(ByVal sendIndex As Integer, UserIndex As Integer)
    On Error Resume Next
    Dim j As Byte

    Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & FONTTYPE_INFO)

    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).OBJIndex Then
            Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).OBJIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount & FONTTYPE_INFO)
        End If
    Next

End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, UserIndex As Integer)
    On Error Resume Next
    Dim j As Integer
    Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
    For j = 1 To NUMSKILLS
        Call SendData(ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & FONTTYPE_INFO)
    Next
End Sub
Sub Parpa(UserIndex As Integer)
    Call SendData(ToIndex, UserIndex, 0, "PARPA")
End Sub
Sub UpdateFuerzaYAg(UserIndex As Integer)
    Dim Fue As Integer
    Dim Agi As Integer

    Fue = UserList(UserIndex).Stats.UserAtributos(fuerza)
    If Fue = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) Then Fue = 0

    Agi = UserList(UserIndex).Stats.UserAtributos(Agilidad)
    If Agi = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) Then Agi = 0

    Call SendData(ToIndex, UserIndex, 0, "EIFYA" & Fue & "," & Agi)

End Sub
Sub UpdateUserMap(UserIndex As Integer)
    On Error GoTo ErrorHandler
    Dim TempChar As Integer
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim i As Integer

    Map = UserList(UserIndex).POS.Map
    Call SendData(ToIndex, UserIndex, 0, "ET")


    For i = 1 To MapInfo(Map).NumUsers
        TempChar = MapInfo(Map).UserIndex(i)
        Call MakeUserChar(ToIndex, UserIndex, 0, TempChar, Map, UserList(TempChar).POS.X, UserList(TempChar).POS.Y)
    Next


    For i = 1 To LastNPC
        If Npclist(i).flags.NPCActive And UserList(UserIndex).POS.Map = Npclist(i).POS.Map Then
            Call MakeNPCChar(ToIndex, UserIndex, 0, i, Map, Npclist(i).POS.X, Npclist(i).POS.Y)
        End If
    Next


    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(Map, X, Y).OBJInfo.OBJIndex Then
                If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType <> OBJTYPE_ARBOLES Or MapData(Map, X, Y).trigger = 2 Then
                    If Y >= 40 Then
                        Y = Y
                    End If

                    Call MakeObj(ToIndex, UserIndex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)

                    If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
                        Call Bloquear(ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                        Call Bloquear(ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                    End If
                End If
            End If
        Next
    Next

    Exit Sub
ErrorHandler:
    Call LogError("Error en el sub.UpdateUserMap. Mapa: " & Map & "-" & X & "-" & Y)

End Sub

Function DameUserindex(SocketId As Integer) As Integer

    Dim LoopC As Integer

    LoopC = 1

    Do Until UserList(LoopC).ConnID = SocketId

        LoopC = LoopC + 1

        If LoopC > MaxUsers Then
            DameUserindex = 0
            Exit Function
        End If

    Loop

    DameUserindex = LoopC

End Function
Function EsMascotaCiudadano(NpcIndex As Integer, UserIndex As Integer) As Boolean

    If Npclist(NpcIndex).MaestroUser Then
        EsMascotaCiudadano = UserList(UserIndex).Faccion.Bando = Real
        If EsMascotaCiudadano Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "F0" & UserList(UserIndex).Name)
    End If

End Function
Function EsMascotaCriminal(NpcIndex As Integer, UserIndex As Integer) As Boolean

    If Npclist(NpcIndex).MaestroUser Then
        EsMascotaCriminal = Not UserList(UserIndex).Faccion.Bando = Caos
        If EsMascotaCriminal Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "F0" & UserList(UserIndex).Name)
    End If

End Function
Sub NpcAtacado(NpcIndex As Integer, UserIndex As Integer)

    Npclist(NpcIndex).flags.AttackedBy = UserIndex

    If Npclist(NpcIndex).MaestroUser Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
    If Npclist(NpcIndex).flags.Faccion <> Neutral Then
        If UserList(UserIndex).Faccion.Ataco(Npclist(NpcIndex).flags.Faccion) = 0 Then UserList(UserIndex).Faccion.Ataco(Npclist(NpcIndex).flags.Faccion) = 2
    End If

    Npclist(NpcIndex).Movement = NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1

End Sub
Function PuedeApuñalar(UserIndex As Integer) As Boolean

    If UserList(UserIndex).Invent.WeaponEqpObjIndex Then PuedeApuñalar = ((UserList(UserIndex).Stats.UserSkills(Apuñalar) >= MIN_APUÑALAR) And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) Or ((UserList(UserIndex).Clase = ASESINO) And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))

End Function
Sub SubirSkill(UserIndex As Integer, Skill As Integer, Optional Prob As Integer)
    On Error GoTo errhandler

    If UserList(UserIndex).flags.Hambre = 1 Or UserList(UserIndex).flags.Sed = 1 Then Exit Sub

    If Prob = 0 Then
        If UserList(UserIndex).Stats.ELV <= 3 Then
            Prob = 2
        ElseIf UserList(UserIndex).Stats.ELV > 3 _
               And UserList(UserIndex).Stats.ELV < 6 Then
            Prob = 2
        ElseIf UserList(UserIndex).Stats.ELV >= 6 _
               And UserList(UserIndex).Stats.ELV < 10 Then
            Prob = 2
        ElseIf UserList(UserIndex).Stats.ELV >= 10 _
               And UserList(UserIndex).Stats.ELV < 20 Then
            Prob = 2
        Else
            Prob = 2
        End If
    End If

    If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

    If Int(RandomNumber(1, Prob)) = 2 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(UserList(UserIndex).Stats.ELV).LevelValue Then
        Call AddtoVar(UserList(UserIndex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
        Call SendData(ToIndex, UserIndex, 0, "G0" & SkillsNames(Skill) & "," & UserList(UserIndex).Stats.UserSkills(Skill))
        Call AddtoVar(UserList(UserIndex).Stats.Exp, 50, MAXEXP)
        Call SendData(ToIndex, UserIndex, 0, "EX" & 50)
        Call SendUserEXP(UserIndex)
        Call CheckUserLevel(UserIndex)
    End If
    Exit Sub

errhandler:
    Call LogError("Error en SubirSkill: " & Err.Description & "-" & UserList(UserIndex).Name & "-" & SkillsNames(Skill))
End Sub
Sub BajarInvisible(UserIndex As Integer)

    If UserList(UserIndex).Stats.ELV >= 34 Or UserList(UserIndex).flags.GolpeoInvi Then
        Call QuitarInvisible(UserIndex)
    Else: UserList(UserIndex).flags.GolpeoInvi = 1
    End If

End Sub
Sub QuitarInvisible(UserIndex As Integer)

    UserList(UserIndex).Counters.Invisibilidad = 0
    UserList(UserIndex).flags.Invisible = 0
    UserList(UserIndex).flags.GolpeoInvi = 0
    UserList(UserIndex).flags.Oculto = 0
    Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, ("V3" & UserList(UserIndex).Char.CharIndex & ",0"))

End Sub
Sub UserDie(UserIndex As Integer)
    On Error GoTo ErrorHandler

    If UserList(UserIndex).flags.death = True Then
        Call death_muere(UserIndex)
    End If

    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_USERMUERTE)

    If UserList(UserIndex).flags.Montado = 1 Then Desmontar (UserIndex)

    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "QDL" & UserList(UserIndex).Char.CharIndex)

    UserList(UserIndex).Stats.MinHP = 0
    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.AtacadoPorUser = 0
    UserList(UserIndex).flags.Envenenado = 0
    UserList(UserIndex).flags.Muerto = 1

    If UserList(UserIndex).flags.EnDuelo = True Then
        UserList(UserIndex).flags.EnDuelo = False
        Call WarpUserChar(UserIndex, 17, 50, 50)
        Call SendData(ToAll, 0, 0, "||Duelos> " & UserList(UserIndex).Name & " ha sido vencido." & FONTTYPE_TALK)
    End If

    If Not EsNewbie(UserIndex) Then
        Call TirarTodosLosItems(UserIndex)
    Else
        Call TirarTodosLosItemsNoNewbies(UserIndex)
    End If

    If TransFacc = True Then
        With UserList(UserIndex)
            If .POS.Map = 1 Or .POS.Map = 34 Or .POS.Map = 59 Or .POS.Map = 45 Or .POS.Map = 44 Or .POS.Map = 43 Or .POS.Map = 42 Or .POS.Map = 41 Or .POS.Map = 40 Then
                If .Faccion.Bando = 0 Then    'Neutro
                    Call WarpUserChar(UserIndex, 172, 38, 30, False)
                ElseIf .Faccion.Bando = 1 Then    'Ciuda
                    Call WarpUserChar(UserIndex, 172, 62, 58, False)
                ElseIf .Faccion.Bando = 2 Then    'Crimi
                    Call WarpUserChar(UserIndex, 172, 30, 83, False)
                End If
            End If
        End With
    End If

    Dim aN As Integer

    aN = UserList(UserIndex).flags.AtacadoPorNpc

    If UserList(UserIndex).POS.Map = 12 Then
        Call WarpUserChar(UserIndex, 1, 50, 50, True)
        Call SendData(ToIndex, UserIndex, 0, "||Has sido teletransportado a Ullathorpe!" & FONTTYPE_TALK)
    End If

    If aN Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = 0
    End If

    If UserList(UserIndex).flags.Paralizado Then
        Call SendData(ToIndex, UserIndex, 0, "P8")
        UserList(UserIndex).flags.Paralizado = 0
    End If

    If UserList(UserIndex).flags.Trabajando Then Call SacarModoTrabajo(UserIndex)

    If UserList(UserIndex).flags.Invisible And UserList(UserIndex).flags.AdminInvisible = 0 Then
        Call QuitarInvisible(UserIndex)
    End If

    If UserList(UserIndex).flags.Ceguera = 1 Then
        UserList(UserIndex).Counters.Ceguera = 0
        UserList(UserIndex).flags.Ceguera = 0
        Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, "NSEGUE")
    End If

    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).Counters.Estupidez = 0
        UserList(UserIndex).flags.Estupidez = 0
        Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, "NESTUP")
    End If

    If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call SendData(ToIndex, UserIndex, 0, "DOK")
    End If

    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call SendData(ToIndex, UserIndex, 0, "MEDOK")
    End If

    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    If UserList(UserIndex).Invent.WeaponEqpObjIndex Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
    If UserList(UserIndex).Invent.CascoEqpObjIndex Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpslot)
    If UserList(UserIndex).Invent.MunicionEqpObjIndex Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)

    If UserList(UserIndex).flags.EstaDueleando = True Then
        Call TerminarDuelo(UserList(UserIndex).flags.Oponente, UserIndex)
    End If

    If UserList(UserIndex).Char.loops = LoopAdEternum Then
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
    End If

    If UserList(UserIndex).flags.automatico = True Then
        Call Rondas_UsuarioMuere(UserIndex)
    End If

    If UserList(UserIndex).flags.Navegando = 0 Then
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    Else
        UserList(UserIndex).Char.Body = iFragataFantasmal
    End If

    Dim i As Integer
    For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(i) Then
            If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia Then
                Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
            Else
                Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
                Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
            End If
        End If

    Next

    If UserList(UserIndex).Reto.Retando_2 = True Then
        If UserList(OPCDuelos.J1).flags.Muerto And UserList(OPCDuelos.J2).flags.Muerto Then    'Pareja 1 = muerta.
            Call SendData(ToAll, 0, 0, "||Ring 2> " & UserList(OPCDuelos.J3).Name & " y " & UserList(OPCDuelos.J4).Name & _
                                     " derrotan a " & UserList(OPCDuelos.J1).Name & " y " & UserList(OPCDuelos.J2).Name & FONTTYPE_TALK)

            'Reseteamos flags.
            UserList(OPCDuelos.J1).Reto.Retando_2 = False
            UserList(OPCDuelos.J1).Reto.Send_Request = False
            UserList(OPCDuelos.J1).Reto.Received_Request = False

            UserList(OPCDuelos.J2).Reto.Retando_2 = False
            UserList(OPCDuelos.J2).Reto.Send_Request = False
            UserList(OPCDuelos.J2).Reto.Received_Request = False

            UserList(OPCDuelos.J3).Reto.Send_Request = False
            UserList(OPCDuelos.J3).Reto.Received_Request = False
            UserList(OPCDuelos.J3).Reto.Retando_2 = False

            UserList(OPCDuelos.J4).Reto.Send_Request = False
            UserList(OPCDuelos.J4).Reto.Received_Request = False
            UserList(OPCDuelos.J4).Reto.Retando_2 = False

            Call WarpUserChar(OPCDuelos.J1, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J2, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J3, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J4, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y + 1)    'los mando a ulla

            OPCDuelos.ParejaEspera = False
            OPCDuelos.OCUP = False

            OPCDuelos.J1 = 0
            OPCDuelos.J2 = 0
            OPCDuelos.J3 = 0
            OPCDuelos.J4 = 0

        ElseIf UserList(OPCDuelos.J3).flags.Muerto And UserList(OPCDuelos.J4).flags.Muerto Then
            Call SendData(ToAll, 0, 0, "||Ring 2> " & UserList(OPCDuelos.J1).Name & " y " & UserList(OPCDuelos.J2).Name & _
                                     " derrotan a " & UserList(OPCDuelos.J3).Name & " y " & UserList(OPCDuelos.J4).Name & FONTTYPE_TALK)

            'Reseteamos flags.
            UserList(OPCDuelos.J1).Reto.Retando_2 = False
            UserList(OPCDuelos.J1).Reto.Send_Request = False
            UserList(OPCDuelos.J1).Reto.Received_Request = False

            UserList(OPCDuelos.J2).Reto.Retando_2 = False
            UserList(OPCDuelos.J2).Reto.Send_Request = False
            UserList(OPCDuelos.J2).Reto.Received_Request = False

            UserList(OPCDuelos.J3).Reto.Send_Request = False
            UserList(OPCDuelos.J3).Reto.Received_Request = False
            UserList(OPCDuelos.J3).Reto.Retando_2 = False

            UserList(OPCDuelos.J4).Reto.Send_Request = False
            UserList(OPCDuelos.J4).Reto.Received_Request = False
            UserList(OPCDuelos.J4).Reto.Retando_2 = False

            Call WarpUserChar(OPCDuelos.J1, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J2, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J3, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J4, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y + 1)    'los mando a ulla

            OPCDuelos.ParejaEspera = False
            OPCDuelos.OCUP = False
            OPCDuelos.J1 = 0
            OPCDuelos.J2 = 0
            OPCDuelos.J3 = 0
            OPCDuelos.J4 = 0

        End If

    End If

    If UserList(UserIndex).POS.Map <> 8 Then UserList(UserIndex).Stats.VecesMurioUsuario = UserList(UserIndex).Stats.VecesMurioUsuario + 1

    UserList(UserIndex).NroMascotas = 0

    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, val(UserIndex), UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    If PuedeDestrabarse(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "||Estás encerrado, para destrabarte presiona la tecla Z." & FONTTYPE_INFO)
    Call SendUserStatsBox(UserIndex)

    Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE")

End Sub
Sub ContarMuerte(Muerto As Integer, Atacante As Integer)
    If EsNewbie(Muerto) Then Exit Sub

    If UserList(Muerto).POS.Map = 8 Then Exit Sub

    If UserList(Atacante).flags.LastMatado(UserList(Muerto).Faccion.Bando) <> UCase$(UserList(Muerto).Name) Then
        UserList(Atacante).flags.LastMatado(UserList(Muerto).Faccion.Bando) = UCase$(UserList(Muerto).Name)
        Call AddtoVar(UserList(Atacante).Faccion.Matados(UserList(Muerto).Faccion.Bando), 1, 65000)
    End If

End Sub

Sub Tilelibre(POS As WorldPos, nPos As WorldPos)


    Dim Notfound As Boolean
    Dim LoopC As Integer
    Dim tX As Integer
    Dim tY As Integer
    Dim hayobj As Boolean
    hayobj = False
    nPos.Map = POS.Map

    Do While Not LegalPos(POS.Map, nPos.X, nPos.Y) Or hayobj

        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If

        For tY = POS.Y - LoopC To POS.Y + LoopC
            For tX = POS.X - LoopC To POS.X + LoopC

                If LegalPos(nPos.Map, tX, tY) Then
                    hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.OBJIndex > 0)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        tX = POS.X + LoopC
                        tY = POS.Y + LoopC
                    End If
                End If

            Next
        Next

        LoopC = LoopC + 1

    Loop

    If Notfound Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub
Sub AgregarAUsersPorMapa(UserIndex As Integer)


    MapInfo(UserList(UserIndex).POS.Map).NumUsers = MapInfo(UserList(UserIndex).POS.Map).NumUsers + 1
    If MapInfo(UserList(UserIndex).POS.Map).NumUsers < 0 Then MapInfo(UserList(UserIndex).POS.Map).NumUsers = 0

    If MapInfo(UserList(UserIndex).POS.Map).NumUsers = 1 Then
        ReDim MapInfo(UserList(UserIndex).POS.Map).UserIndex(1 To 1)
    Else

        ReDim Preserve MapInfo(UserList(UserIndex).POS.Map).UserIndex(1 To MapInfo(UserList(UserIndex).POS.Map).NumUsers)
    End If


    MapInfo(UserList(UserIndex).POS.Map).UserIndex(MapInfo(UserList(UserIndex).POS.Map).NumUsers) = UserIndex

End Sub
Sub QuitarDeUsersPorMapa(UserIndex As Integer)


    MapInfo(UserList(UserIndex).POS.Map).NumUsers = MapInfo(UserList(UserIndex).POS.Map).NumUsers - 1
    If MapInfo(UserList(UserIndex).POS.Map).NumUsers < 0 Then MapInfo(UserList(UserIndex).POS.Map).NumUsers = 0

    If MapInfo(UserList(UserIndex).POS.Map).NumUsers Then
        Dim i As Integer

        For i = 1 To MapInfo(UserList(UserIndex).POS.Map).NumUsers + 1

            If MapInfo(UserList(UserIndex).POS.Map).UserIndex(i) = UserIndex Then Exit For
        Next

        For i = i To MapInfo(UserList(UserIndex).POS.Map).NumUsers

            MapInfo(UserList(UserIndex).POS.Map).UserIndex(i) = MapInfo(UserList(UserIndex).POS.Map).UserIndex(i + 1)
        Next

        ReDim Preserve MapInfo(UserList(UserIndex).POS.Map).UserIndex(1 To MapInfo(UserList(UserIndex).POS.Map).NumUsers)
    Else
        ReDim MapInfo(UserList(UserIndex).POS.Map).UserIndex(0)
    End If

End Sub
Sub WarpUserChar(UserIndex As Integer, Map As Integer, X As Integer, Y As Integer, Optional FX As Boolean = False)

    Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, "QDL" & UserList(UserIndex).Char.CharIndex)
    Call SendData(ToIndex, UserIndex, UserList(UserIndex).POS.Map, "QTDL")

    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer

    UserList(UserIndex).Counters.Protegido = 2
    UserList(UserIndex).flags.Protegido = 3

    OldMap = UserList(UserIndex).POS.Map
    OldX = UserList(UserIndex).POS.X
    OldY = UserList(UserIndex).POS.Y

    Call EraseUserChar(ToMap, 0, OldMap, UserIndex)

    UserList(UserIndex).POS.X = X
    UserList(UserIndex).POS.Y = Y

    If OldMap = Map Then
        Call MakeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y)
        Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)
    Else
        Call QuitarDeUsersPorMapa(UserIndex)
        UserList(UserIndex).POS.Map = Map
        Call AgregarAUsersPorMapa(UserIndex)

        Call SendData(ToIndex, UserIndex, 0, "CM" & UserList(UserIndex).POS.Map & "," & MapInfo(UserList(UserIndex).POS.Map).MapVersion & "," & MapInfo(UserList(UserIndex).POS.Map).Name & "," & MapInfo(UserList(UserIndex).POS.Map).TopPunto & "," & MapInfo(UserList(UserIndex).POS.Map).LeftPunto)
        If MapInfo(Map).Music <> MapInfo(OldMap).Music Then Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)

        Call MakeUserChar(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y)
        Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)
    End If

    Call UpdateUserMap(UserIndex)

    If FX And UserList(UserIndex).flags.AdminInvisible = 0 And Not UserList(UserIndex).flags.Meditando Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_WARP)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
    End If
    Dim i As Integer

    For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(i) Then
            If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
            End If
        End If
    Next

End Sub
Sub WarpMascotas(UserIndex As Integer)
    Dim i As Integer

    Dim UMascRespawn As Boolean
    Dim miflag As Byte, MascotasReales As Integer
    Dim prevMacotaType As Integer

    Dim PetTypes(1 To MAXMASCOTAS) As Integer
    Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
    Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

    Dim NroPets As Integer

    NroPets = UserList(UserIndex).NroMascotas

    For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(i) Then
            PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
            If PetRespawn(i) Then
                PetTypes(i) = UserList(UserIndex).MascotasType(i)
                PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
            Else
                PetTypes(i) = UserList(UserIndex).MascotasType(i)
                PetTiempoDeVida(i) = 1
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
            End If
        End If
    Next

    For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If PetTypes(i) Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).POS, False, PetRespawn(i))
            UserList(UserIndex).MascotasType(i) = PetTypes(i)

            If UserList(UserIndex).MascotasIndex(i) = MAXNPCS Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
                If UserList(UserIndex).NroMascotas Then UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
                Exit Sub
            End If
            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = SIGUE_AMO
            Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNpc = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call QuitarNPCDeLista(Npclist(UserList(UserIndex).MascotasIndex(i)).Numero, UserList(UserIndex).POS.Map)
            Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
        End If
    Next

    UserList(UserIndex).NroMascotas = NroPets

End Sub
Sub Cerrar_Usuario(UserIndex As Integer)

    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True
        UserList(UserIndex).Counters.Salir = Timer - 8 * Buleano(UserList(UserIndex).Clase = PIRATA And UserList(UserIndex).Recompensas(3) = 2)
        Call SendData(ToIndex, UserIndex, 0, "1Z" & IntervaloCerrarConexion - 8 * Buleano(UserList(UserIndex).Clase = PIRATA And UserList(UserIndex).Recompensas(3) = 2))
    End If

End Sub
