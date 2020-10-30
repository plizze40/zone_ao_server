Attribute VB_Name = "Handledata_1"
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
Public Sub HandleData1(UserIndex As Integer, rdata As String, Procesado As Boolean)
    Dim tInt As Integer, TIndex As Integer, X As Integer, Y As Integer
    Dim Arg1 As String, Arg2 As String, arg3 As String
    Dim nPos As WorldPos
    Dim tLong As Long
    Dim ind

    Procesado = True

    Select Case UCase$(Left$(rdata, 1))
    Case "\"
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        rdata = Right$(rdata, Len(rdata) - 1)
        tName = ReadField(1, rdata, 32)
        TIndex = NameIndex(tName)

        If TIndex <> 0 Then
            If UserList(TIndex).flags.Muerto = 1 Then Exit Sub

            If Len(rdata) <> Len(tName) Then
                tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
            Else
                tMessage = " "
            End If

            If Not EnPantalla(UserList(UserIndex).POS, UserList(TIndex).POS, 1) Then
                Call SendData(ToIndex, UserIndex, 0, "2E")
                Exit Sub
            End If

            ind = UserList(UserIndex).Char.CharIndex

            If InStr(tMessage, "°") Then Exit Sub

            If UserList(TIndex).flags.Privilegios > 0 And UserList(UserIndex).flags.Privilegios = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "3E")
                Exit Sub
            End If

            Call SendData(ToIndex, UserIndex, UserList(UserIndex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Call SendData(ToIndex, TIndex, UserList(UserIndex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Call SendData(ToGMArea, UserIndex, UserList(UserIndex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Exit Sub
        End If

        Call SendData(ToIndex, UserIndex, 0, "3E")
        Exit Sub

    Case ";"
        Dim Modo As String

        rdata = Right$(rdata, Len(rdata) - 1)
        If Right$(rdata, Len(rdata) - 1) = " " Or Right$(rdata, Len(rdata) - 1) = "-" Then rdata = "1 "
        If Len(rdata) = 1 Then Exit Sub

        If val(Right$(rdata, Len(rdata) - 1)) > 0 Then
            'If UserList(UserIndex).flags.Trabajando Then Exit Sub
            'UserList(UserIndex).flags.IntentosCodigo = UserList(UserIndex).flags.IntentosCodigo + 1
            'If UserList(UserIndex).flags.IntentosCodigo >= 10 Then
            'UserList(UserIndex).flags.CodigoTrabajo = 0
            'UserList(UserIndex).flags.IntentosCodigo = 0
            'Call SacarModoTrabajo(UserIndex)
            'Call SendData(ToIndex, UserIndex, 0, "||Fuiste encarcelado por mandar demasiados códigos de trabajo incorrectos." & FONTTYPE_FIGHT)
            'Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " fue encarcelado por mandar demasiados códigos incorrectos." & fonttype_Server)
            'Call Encarcelar(UserIndex, 15)
            'Exit Sub
            'Else: Call SendData(ToIndex, UserIndex, 0, "||Código incorrecto. Te quedan " & 10 - UserList(UserIndex).flags.IntentosCodigo & " intentos o serás encarcelado." & FONTTYPE_INFO)
            'End If
        End If


        Modo = Left$(rdata, 1)
        rdata = Replace(Right$(rdata, Len(rdata) - 1), "~", "-")

        Select Case Modo

        Case 1, 4, 5, 8, 9

            If InStr(rdata, "°") Then Exit Sub

            If (Modo = 4 Or Modo = 5) And UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If

            If UserList(UserIndex).flags.Privilegios = 1 Or UserList(UserIndex).flags.Privilegios = 2 Or UserList(UserIndex).flags.Privilegios = 3 Then Call LogGM(UserList(UserIndex).Name, "Dijo: " & rdata, True)
            If InStr(1, rdata, Chr$(255)) Then rdata = Replace(rdata, Chr$(255), " ")

            ind = UserList(UserIndex).Char.CharIndex
            Dim Color As Long
            Dim IndexSendData As Byte

            If Modo = 4 Then
                Color = vbRed

            ElseIf Modo = 5 Then
                Color = vbGreen
            ElseIf UserList(UserIndex).flags.Privilegios Then
                Color = &H80FF&
            ElseIf UserList(UserIndex).flags.EsConseCaos And UserList(UserIndex).Faccion.Bando = Caos Then    'Real
                Color = &H40C0&
            ElseIf UserList(UserIndex).flags.EsConseReal And UserList(UserIndex).Faccion.Bando = Real Then  'Caos
                Color = &HC0C000
            ElseIf UserList(UserIndex).flags.EsConcilioNegro And UserList(UserIndex).Faccion.Bando = Neutral Then  'Caos
                Color = &H808080
            ElseIf UserList(UserIndex).flags.Quest And UserList(UserIndex).Faccion.Bando <> Neutral Then
                If UserList(UserIndex).Faccion.Bando = Real Then
                    Color = vbBlue
                Else: Color = vbRed
                End If

            ElseIf UserList(UserIndex).Stats.ELV = 45 Then
                Color = &HFF8080

            ElseIf UserList(UserIndex).flags.Muerto Then
                Color = vbYellow
            Else: Color = vbWhite
            End If

            If UserList(UserIndex).flags.Privilegios > 0 Or UserList(UserIndex).Clase = CLERIGO Then
                IndexSendData = ToPCArea
            ElseIf UserList(UserIndex).flags.Muerto Then
                IndexSendData = ToMuertos
            Else
                IndexSendData = ToPCAreaVivos
            End If

            If UCase$(rdata) = "SACRIFICATE!" Then
                nPos = UserList(UserIndex).POS
                Call HeadtoPos(UserList(UserIndex).Char.Heading, nPos)
                TIndex = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                If TIndex > 0 Then
                    If MapData(nPos.Map, nPos.X - 1, nPos.Y).OBJInfo.OBJIndex = Cruz And _
                       MapData(nPos.Map, nPos.X + 1, nPos.Y).OBJInfo.OBJIndex = Cruz And _
                       MapData(nPos.Map, nPos.X, nPos.Y - 1).OBJInfo.OBJIndex = Cruz And _
                       MapData(nPos.Map, nPos.X, nPos.Y + 1).OBJInfo.OBJIndex = Cruz Then
                        If UserList(UserIndex).Stats.ELV < 40 Then
                            Call SendData(ToIndex, UserIndex, 0, "||Debes ser nivel 40 o más para iniciar un sacrificio." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If UserList(TIndex).Stats.MinHP < UserList(TIndex).Stats.MaxHP / 2 Then
                            Call SendData(ToIndex, UserIndex, 0, "||Solo puedes comenzar a sacrificar a usuarios que tengan más de la mitad de sus HP." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        UserList(TIndex).flags.Sacrificando = 1
                        UserList(TIndex).flags.Sacrificador = UserIndex
                        UserList(UserIndex).flags.Sacrificado = TIndex
                        Call SendData(ToIndex, UserIndex, 0, "||¡Comenzaste a sacrificar a " & UserList(TIndex).Name & "!" & FONTTYPE_INFO)
                        Call SendData(ToIndex, TIndex, 0, "||¡" & UserList(UserIndex).Name & " comenzó a sacrificarte! ¡Huye!" & FONTTYPE_INFO)
                    End If
                End If
            End If

            If Modo = 5 Then rdata = "* " & rdata & " *"

            Call SendData(IndexSendData, UserIndex, UserList(UserIndex).POS.Map, "||" & Color & "°" & rdata & "°" & str(ind))
            Exit Sub

        Case 2

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If

            TIndex = UserList(UserIndex).flags.Whispereando

            If TIndex Then
                If UserList(TIndex).flags.Muerto Then Exit Sub

                If Not EnPantalla(UserList(UserIndex).POS, UserList(TIndex).POS, 1) Then
                    Call SendData(ToIndex, UserIndex, 0, "2E")
                    Exit Sub
                End If

                ind = UserList(UserIndex).Char.CharIndex

                If InStr(rdata, "°") Then Exit Sub

                If UserList(TIndex).flags.Privilegios > 0 And UserList(TIndex).flags.AdminInvisible Then
                    Call SendData(ToIndex, UserIndex, 0, "3E")
                    Call SendData(ToIndex, TIndex, UserList(UserIndex).POS.Map, "||" & vbBlue & "°" & rdata & "°" & str(ind))
                    Exit Sub
                End If

                If UserList(UserIndex).flags.Privilegios = 1 Then Call LogGM(UserList(UserIndex).Name, "Grito: " & rdata, True)

                If EnPantalla(UserList(UserIndex).POS, UserList(TIndex).POS, 1) Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                    Call SendData(ToIndex, TIndex, 0, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                    Call SendData(ToGMArea, UserIndex, UserList(UserIndex).POS.Map, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                Else
                    Call SendData(ToIndex, UserIndex, 0, "{F")
                    UserList(UserIndex).flags.Whispereando = 0
                End If
            End If

            Exit Sub

        Case 3
            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If

            If Len(rdata) And Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then Call SendData(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_GUILD)
            Exit Sub

        Case 6
            If UserList(UserIndex).flags.Party = 0 Then Exit Sub

            If Len(rdata) > 0 Then
                Call SendData(ToParty, UserIndex, 0, "||" & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_PARTY)
            End If
            Exit Sub

        Case 7
            If UserList(UserIndex).flags.Privilegios = 0 Then Exit Sub

            Call LogGM(UserList(UserIndex).Name, "Mensaje a Gms:" & rdata, (UserList(UserIndex).flags.Privilegios = 1))
            If Len(rdata) > 0 Then
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & "> " & rdata & "~255~255~255~0~1")
                Call SendData(ToConse, 0, 0, "||" & UserList(UserIndex).Name & "> " & rdata & "~255~255~255~0~1")
                Call SendData(ToConci, 0, 0, "||" & UserList(UserIndex).Name & "> " & rdata & "~255~255~255~0~1")
            End If

            Exit Sub

        End Select

    Case "M"
        Dim Mide As Double
        rdata = Right$(rdata, Len(rdata) - 1)

        If UserList(UserIndex).flags.Trabajando Then

            Call SacarModoTrabajo(UserIndex)

        End If

        If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando _
           And UserList(UserIndex).flags.Paralizado = 0 Then
            Call MoveUserChar(UserIndex, val(rdata))
        ElseIf UserList(UserIndex).flags.Descansar Then
            UserList(UserIndex).flags.Descansar = False
            Call SendData(ToIndex, UserIndex, 0, "DOK")
            Call SendData(ToIndex, UserIndex, 0, "DN")
            Call MoveUserChar(UserIndex, val(rdata))
        End If

        If UserList(UserIndex).flags.Oculto Then
            If Not (UserList(UserIndex).Clase = LADRON And UserList(UserIndex).Recompensas(2) = 1) Then
                UserList(UserIndex).flags.Oculto = 0
                UserList(UserIndex).flags.Invisible = 0
                Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, ("V3" & UserList(UserIndex).Char.CharIndex & ",0"))
                Call SendData(ToIndex, UserIndex, 0, "V5")
            End If
        End If

        Exit Sub
    End Select

    Select Case UCase$(Left$(rdata, 2))
    Case "ZI"
        rdata = Right$(rdata, Len(rdata) - 2)
        Dim Bait(1 To 2) As Byte
        Bait(1) = val(ReadField(1, rdata, 44))
        Bait(2) = val(ReadField(2, rdata, 44))

        Select Case Bait(2)
        Case 0
            Bait(2) = Bait(1) - 1
        Case 1
            Bait(2) = Bait(1) + 1
        Case 2
            Bait(2) = Bait(1) - 5
        Case 3
            Bait(2) = Bait(1) + 5
        End Select

        If Bait(2) > 0 And Bait(2) <= MAX_INVENTORY_SLOTS Then Call AcomodarItems(UserIndex, Bait(1), Bait(2))

        Exit Sub
    Case "TI"
        If UserList(UserIndex).flags.Navegando = 1 Or _
           UserList(UserIndex).flags.Muerto = 1 Or _
           UserList(UserIndex).flags.Montado Then Exit Sub


        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If val(Arg1) = FLAGORO Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes tirar oro." & FONTTYPE_WARNING)
            Exit Sub
        Else
            If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) Then
                If UserList(UserIndex).Invent.Object(val(Arg1)).OBJIndex = 0 Then
                    Exit Sub
                End If
                Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y)
            Else
                Exit Sub
            End If
        End If
        Exit Sub
    Case "SF"
        rdata = Right$(rdata, Len(rdata) - 2)
        If Not PuedeFaccion(UserIndex) Then Exit Sub
        If UserList(UserIndex).Faccion.BandoOriginal Then Exit Sub
        tInt = val(rdata)

        If tInt = Neutral Then
            If UserList(UserIndex).Faccion.Bando <> Neutral Then
                Call SendData(ToIndex, UserIndex, 0, "7&")
            Else: Call SendData(ToIndex, UserIndex, 0, "0&")
            End If
            Exit Sub
        End If

        If UserList(UserIndex).Faccion.Matados(tInt) > UserList(UserIndex).Faccion.Matados(Enemigo(tInt)) Then
            Call SendData(ToIndex, UserIndex, 0, Mensajes(tInt, 9))
            Exit Sub
        End If

        Call SendData(ToIndex, UserIndex, 0, Mensajes(tInt, 10))
        UserList(UserIndex).Faccion.BandoOriginal = tInt
        UserList(UserIndex).Faccion.Bando = tInt
        UserList(UserIndex).Faccion.Ataco(tInt) = 0
        If Not PuedeFaccion(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SUFA0")

        Call UpdateUserChar(UserIndex)

        Exit Sub
    Case "LH"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 2)
        UserList(UserIndex).flags.Hechizo = val(rdata)
        Exit Sub
    Case "WH"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        If Not InMapBounds(X, Y) Then Exit Sub
        Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, X, Y)

        If UserList(UserIndex).flags.TargetUser = UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "{C")
            Exit Sub
        End If

        If UserList(UserIndex).flags.TargetUser Then
            UserList(UserIndex).flags.Whispereando = UserList(UserIndex).flags.TargetUser
            Call SendData(ToIndex, UserIndex, 0, "{B" & UserList(UserList(UserIndex).flags.Whispereando).Name)
        Else
            Call SendData(ToIndex, UserIndex, 0, "{D")
        End If

        Exit Sub
    Case "LC"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        Dim POS As WorldPos
        POS.Map = UserList(UserIndex).POS.Map
        POS.X = CInt(Arg1)
        POS.Y = CInt(Arg2)
        If Not EnPantalla(UserList(UserIndex).POS, POS, 1) Then Exit Sub
        Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, POS.X, POS.Y)
        Exit Sub
    Case "RC"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        Call Accion(UserIndex, UserList(UserIndex).POS.Map, X, Y)
        Exit Sub
    Case "UK"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        rdata = Right$(rdata, Len(rdata) - 2)
        Select Case val(rdata)
        Case Robar
            Call SendData(ToIndex, UserIndex, 0, "T01" & Robar)
        Case Magia
            Call SendData(ToIndex, UserIndex, 0, "T01" & Magia)
        Case Domar
            Call SendData(ToIndex, UserIndex, 0, "T01" & Domar)
        Case Invitar
            Call SendData(ToIndex, UserIndex, 0, "T01" & Invitar)

        Case Ocultarse

            If UserList(UserIndex).flags.Navegando Then
                Call SendData(ToIndex, UserIndex, 0, "6E")
                Exit Sub
            End If

            If UserList(UserIndex).flags.Oculto Then
                Call SendData(ToIndex, UserIndex, 0, "7E")
                Exit Sub
            End If

            Call DoOcultarse(UserIndex)
        End Select
        Exit Sub
    End Select

    Select Case UCase$(rdata)
    Case "KLA"
        Call SendData(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & "(" & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & ") pide ayuda en " & UserList(UserIndex).POS.Map & "," & UserList(UserIndex).POS.X & "," & UserList(UserIndex).POS.Y & "" & FONTTYPE_FENIX)
        Exit Sub
    Case "VSI"    'matute - encuesta votos positivos
        If encuestas.activa = 0 Then Exit Sub
        If UserList(UserIndex).flags.votoencuesta = 1 Then Exit Sub
        encuestas.votosSI = encuestas.votosSI + 1
        Call SendData(ToIndex, UserIndex, 0, "||Voto computado." & FONTTYPE_INFO)
        UserList(UserIndex).flags.votoencuesta = 1
        Exit Sub
    Case "VNO"    'matute - encuesta votos negativos
        If encuestas.activa = 0 Then Exit Sub
        If UserList(UserIndex).flags.votoencuesta = 1 Then Exit Sub
        encuestas.votosNP = encuestas.votosNP + 1
        Call SendData(ToIndex, UserIndex, 0, "||Voto computado." & FONTTYPE_INFO)
        UserList(UserIndex).flags.votoencuesta = 1
        Exit Sub
    Case "RPU"
        Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).POS.X & "," & UserList(UserIndex).POS.Y)
        Exit Sub
    Case "AT"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        If UserList(UserIndex).Invent.WeaponEqpObjIndex Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil Or ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes usar así esta arma." & FONTTYPE_FIGHT)
                Exit Sub
            End If
        End If

        Call UsuarioAtaca(UserIndex)

        Exit Sub
    Case "AG"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If





        Call GetObj(UserIndex)
        Exit Sub
    Case "SEG"
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(ToIndex, UserIndex, 0, "1O")
        Else
            Call SendData(ToIndex, UserIndex, 0, "9K")
        End If
        UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
        Exit Sub
    Case "ATRI"
        Call EnviarAtrib(UserIndex)
        Exit Sub
    Case "FAMA"
        Call EnviarFama(UserIndex)
        Call EnviarMiniSt(UserIndex)
        Exit Sub
    Case "ESKI"
        Call EnviarSkills(UserIndex)
        Exit Sub
    Case "PARSAL"
        Dim i As Integer
        If UserList(UserIndex).flags.Party Then
            If Party(UserList(UserIndex).PartyIndex).NroMiembros = 2 Then
                Call RomperParty(UserIndex)
            Else: Call SacarDelParty(UserIndex)
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No estás en party." & FONTTYPE_PARTY)
        End If
        Exit Sub
    Case "PARINF"
        Call EnviarIntegrantesParty(UserIndex)
        Exit Sub

    Case "FINCOM"

        UserList(UserIndex).flags.Comerciando = False
        Call SendData(ToIndex, UserIndex, 0, "FINCOMOK")
        Exit Sub
    Case "FINCOMUSU"
        If UserList(UserIndex).ComUsu.DestUsu > 0 Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "6R" & UserList(UserIndex).Name)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If

        Call FinComerciarUsu(UserIndex)
        Exit Sub

    Case "FINBAN"
        UserList(UserIndex).flags.Comerciando = False
        Call SendData(ToIndex, UserIndex, 0, "FINBANOK")
        Exit Sub

    Case "FINTIE"
        UserList(UserIndex).flags.Comerciando = False
        Call SendData(ToIndex, UserIndex, 0, "FINTIEOK")
        Exit Sub

    Case "COMUSUOK"

        Call AceptarComercioUsu(UserIndex)
        Exit Sub
    Case "COMUSUNO"

        If UserList(UserIndex).ComUsu.DestUsu Then
            Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "7R" & UserList(UserIndex).Name)
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
        End If
        Call SendData(ToIndex, UserIndex, 0, "8R")
        Call FinComerciarUsu(UserIndex)
        Exit Sub

    Case "GLINFO"
        If UserList(UserIndex).GuildInfo.EsGuildLeader Then
            If UserList(UserIndex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, UserIndex, 0, "GINFIG")
            Else
                Call SendGuildLeaderInfo(UserIndex)
            End If
        ElseIf Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then
            If UserList(UserIndex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, UserIndex, 0, "GINFII")
            Else
                Call SendGuildsStats(UserIndex)
            End If
        Else
            If UserList(UserIndex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, UserIndex, 0, "GINFIJ")
            Else: Call SendGuildsList(UserIndex)
            End If
        End If

        Exit Sub

    End Select

    Select Case UCase$(Left$(rdata, 2))
    Case "(A"
        If PuedeDestrabarse(UserIndex) Then
            Call ClosestLegalPos(UserList(UserIndex).POS, nPos)
            If InMapBounds(nPos.X, nPos.Y) Then Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
        End If

        Exit Sub
    Case "GM"
        rdata = Right$(rdata, Len(rdata) - 2)
        Call SendData(ToAdmins, 0, 0, "||Soporte " & "(" & ReadField(2, rdata, Asc("¬")) & ") >" & UserList(UserIndex).Name & " dice: " & ReadField(3, rdata, Asc("¬")) & FONTTYPE_FENIX)
        Exit Sub

    End Select

    Select Case UCase$(Left$(rdata, 3))
    Case "FRF"
        rdata = Right$(rdata, Len(rdata) - 3)
        For i = 1 To 10
            If UserList(UserIndex).flags.Espiado(i) > 0 Then
                If UserList(UserList(UserIndex).flags.Espiado(i)).flags.Privilegios > 1 Then Call SendData(ToIndex, UserList(UserIndex).flags.Espiado(i), 0, "{{" & UserList(UserIndex).Name & "," & rdata)
            End If
        Next
        Exit Sub
    Case "USA"
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
            If UserList(UserIndex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(UserIndex, val(rdata), 0)
        Exit Sub
    Case "USE"
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
            If UserList(UserIndex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(UserIndex, val(rdata), 1)
        Exit Sub
    Case "CNS"
        Dim Arg5 As Integer
        rdata = Right$(rdata, Len(rdata) - 3)

        X = CInt(ReadField(1, rdata, 32))
        Arg5 = CInt(ReadField(2, rdata, 32))
        If Arg5 < 1 Then Exit Sub
        If X < 1 Then Exit Sub
        If ObjData(X).SkHerreria = 0 Then Exit Sub
        Call HerreroConstruirItem(UserIndex, X, val(Arg5))
        Exit Sub

    Case "CNC"
        rdata = Right$(rdata, Len(rdata) - 3)

        X = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If Arg1 < 1 Then Exit Sub
        If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
        Call CarpinteroConstruirItem(UserIndex, X, val(Arg1))
        Exit Sub
    Case "SCR"
        rdata = Right$(rdata, Len(rdata) - 3)

        X = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If X < 1 Or ObjData(X).SkSastreria = 0 Then Exit Sub
        Call SastreConstruirItem(UserIndex, X, val(Arg1))
        Exit Sub

    Case "WLC"
        rdata = Right$(rdata, Len(rdata) - 3)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        arg3 = ReadField(3, rdata, 44)
        If Len(arg3) = 0 Or Len(Arg2) = 0 Or Len(Arg1) = 0 Then Exit Sub
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(arg3) Then Exit Sub

        POS.Map = UserList(UserIndex).POS.Map
        POS.X = CInt(Arg1)
        POS.Y = CInt(Arg2)
        tLong = CInt(arg3)

        If UserList(UserIndex).flags.Muerto = 1 Or _
           UserList(UserIndex).flags.Descansar Or _
           UserList(UserIndex).flags.Meditando Or _
           Not InMapBounds(POS.X, POS.Y) Then Exit Sub

        If Not EnPantalla(UserList(UserIndex).POS, POS, 1) Then
            Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).POS.X & "," & UserList(UserIndex).POS.Y)
            Exit Sub
        End If

        Select Case tLong

        Case Proyectiles
            Dim TU As Integer, tN As Integer

            If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Or _
               UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then Exit Sub

            If UserList(UserIndex).Invent.WeaponEqpSlot < 1 Or UserList(UserIndex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Or _
               UserList(UserIndex).Invent.MunicionEqpSlot < 1 Or UserList(UserIndex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Or _
               ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).ObjType <> OBJTYPE_FLECHAS Or _
               UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).Amount < 1 Or _
               ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then Exit Sub

            If TiempoTranscurrido(UserList(UserIndex).Counters.LastFlecha) < IntervaloUserFlechas Then Exit Sub
            If TiempoTranscurrido(UserList(UserIndex).Counters.LastHechizo) < IntervaloUserPuedeHechiGolpe Then Exit Sub
            If TiempoTranscurrido(UserList(UserIndex).Counters.LastGolpe) < IntervaloUserPuedeAtacar Then Exit Sub

            UserList(UserIndex).Counters.LastFlecha = Timer
            Call SendData(ToIndex, UserIndex, 0, "LF")

            If UserList(UserIndex).Stats.MinSta >= 10 Then
                Call QuitarSta(UserIndex, RandomNumber(1, 10))
            Else
                Call SendData(ToIndex, UserIndex, 0, "9E")
                Exit Sub
            End If

            Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, val(Arg1), val(Arg2))

            TU = UserList(UserIndex).flags.TargetUser
            tN = UserList(UserIndex).flags.TargetNpc

            If TU = UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "3N")
                Exit Sub
            End If

            Call QuitarUnItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)

            If UserList(UserIndex).Invent.MunicionEqpSlot Then
                UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).Equipped = 1
                Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
            End If

            If tN Then
                If Npclist(tN).Attackable Then Call UsuarioAtacaNpc(UserIndex, tN)
            ElseIf TU Then
                If TU <> UserIndex Then
                    Call UsuarioAtacaUsuario(UserIndex, TU)
                    SendUserHP TU
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If






        Case Invitar
            Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, POS.X, POS.Y)

            If UserList(UserIndex).flags.TargetUser = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||No hay nadie a quien invitar." & FONTTYPE_PARTY)
                Exit Sub
            End If

            If UserList(UserIndex).flags.Privilegios > 0 Or UserList(UserList(UserIndex).flags.TargetUser).flags.Privilegios > 0 Then Exit Sub

            Call DoInvitar(UserIndex, UserList(UserIndex).flags.TargetUser)

        Case Magia


            If UserList(UserIndex).flags.Privilegios = 1 Then Exit Sub

            Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, POS.X, POS.Y)

            If UserList(UserIndex).flags.Hechizo Then
                Call LanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                UserList(UserIndex).flags.Hechizo = 0
            Else
                Call SendData(ToIndex, UserIndex, 0, "4N")
            End If

        Case Robar
            If TiempoTranscurrido(UserList(UserIndex).Counters.LastTrabajo) < 1 Then Exit Sub
            If MapInfo(UserList(UserIndex).POS.Map).Pk Or (UserList(UserIndex).Clase = LADRON) Then


                Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, POS.X, POS.Y)

                If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then
                    If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                        nPos.Map = UserList(UserIndex).POS.Map
                        nPos.X = POS.X
                        nPos.Y = POS.Y

                        If Distancia(nPos, UserList(UserIndex).POS) > 4 Or (Not (UserList(UserIndex).Clase = LADRON And UserList(UserIndex).Recompensas(3) = 1) And Distancia(nPos, UserList(UserIndex).POS) > 2) Then
                            Call SendData(ToIndex, UserIndex, 0, "DL")
                            Exit Sub
                        End If

                        Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "4S")
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "5S")
            End If

        Case Domar



            Dim CI As Integer

            Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, POS.X, POS.Y)
            CI = UserList(UserIndex).flags.TargetNpc

            If CI Then
                If Npclist(CI).flags.Domable Then
                    nPos.Map = UserList(UserIndex).POS.Map
                    nPos.X = POS.X
                    nPos.Y = POS.Y
                    If Distancia(nPos, Npclist(UserList(UserIndex).flags.TargetNpc).POS) > 2 Then
                        Call SendData(ToIndex, UserIndex, 0, "DL")
                        Exit Sub
                    End If
                    If Npclist(CI).flags.AttackedBy Then
                        Call SendData(ToIndex, UserIndex, 0, "7S")
                        Exit Sub
                    End If
                    Call DoDomar(UserIndex, CI)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "8S")
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "9S")
            End If

        Case FundirMetal
            Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, POS.X, POS.Y)

            If UserList(UserIndex).flags.TargetObj Then
                If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                    Call FundirMineral(UserIndex)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "8N")
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "8N")
            End If

        Case Herreria
            Call LookatTile(UserIndex, UserList(UserIndex).POS.Map, POS.X, POS.Y)

            If UserList(UserIndex).flags.TargetObj Then
                If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
                    Call EnviarArmasConstruibles(UserIndex)
                    Call EnviarArmadurasConstruibles(UserIndex)
                    Call EnviarEscudosConstruibles(UserIndex)
                    Call EnviarCascosConstruibles(UserIndex)
                    Call SendData(ToIndex, UserIndex, 0, "SFH")
                    UserList(UserIndex).flags.EnviarHerreria = 1
                Else
                    Call SendData(ToIndex, UserIndex, 0, "2T")
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "2T")
            End If
        Case Else

            If UserList(UserIndex).flags.Trabajando = 0 Then
                Dim TrabajoPos As WorldPos
                TrabajoPos.Map = UserList(UserIndex).POS.Map
                TrabajoPos.X = POS.X
                TrabajoPos.Y = POS.Y
                Call InicioTrabajo(UserIndex, tLong, TrabajoPos)
            End If
            Exit Sub

        End Select

        UserList(UserIndex).Counters.LastTrabajo = Timer
        Exit Sub
    Case "REL"
        If UserList(UserIndex).flags.Muerto Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 3)
        Call RecibirRecompensa(UserIndex, val(rdata))
        Exit Sub
    Case "CIG"
        rdata = Right$(rdata, Len(rdata) - 3)
        X = Guilds.Count

        If CreateGuild(UserList(UserIndex).Name, UserIndex, rdata) Then
            If X = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "3T")
            Else
                Call SendData(ToIndex, UserIndex, 0, "4T" & X)
            End If
            Call UpdateUserChar(UserIndex)

        End If

        Exit Sub
    Case "RSB"
        If UserList(UserIndex).flags.Muerto Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 3)
        Call RecibirSubclase(CByte(rdata), UserIndex)
        Exit Sub
    End Select

    Select Case UCase$(Left$(rdata, 4))

    Case "PPCC"
        Dim caption As String
        rdata = Right$(rdata, Len(rdata) - 4)
        caption = ReadField(1, rdata, 44)
        TIndex = ReadField(2, rdata, 44)
        Call SendData(ToIndex, TIndex, 0, "PCCC" & caption & "," & UserList(UserIndex).Name)
        Exit Sub
    Case "PRCS"
        rdata = Right$(rdata, Len(rdata) - 4)
        Call SendData(ToIndex, UserList(UserIndex).flags.EsperandoLista, 0, "PRAP" & rdata)
        If rdata = "@*|" Then UserList(UserIndex).flags.EsperandoLista = 0
        Exit Sub
    Case "PASS"
        rdata = Right$(rdata, Len(rdata) - 4)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)

        If UserList(UserIndex).PassWord <> Arg1 Then
            Call SendData(ToIndex, UserIndex, 0, "||El password viejo provisto no es correcto." & FONTTYPE_INFO)
            Exit Sub
        End If

        UserList(UserIndex).PassWord = Arg2
        Call SendData(ToIndex, UserIndex, 0, "3V")

        Exit Sub
    Case "%PR%"
        rdata = Right$(rdata, Len(rdata) - 4)
        TIndex = ReadField(1, rdata, 32)
        rdata = ReadField(2, rdata, 32)
        Call SendData(ToIndex, TIndex, 0, "||Procesos de " & UCase$(UserList(UserIndex).Name) & ":" & FONTTYPE_VENENO)
        Call SendData(ToIndex, TIndex, 0, "||" & rdata & FONTTYPE_INFO)
        Exit Sub
    Case "INFS"
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
            Dim H As Integer
            H = UserList(UserIndex).Stats.UserHechizos(val(rdata))
            If H > 0 And H < NumeroHechizos + 1 Then
                Call SendData(ToIndex, UserIndex, 0, "7T" & Hechizos(H).Nombre & "¬" & Hechizos(H).Desc & "¬" & Hechizos(H).MinSkill & "¬" & ManaHechizo(UserIndex, H) & "¬" & Hechizos(H).StaRequerido)
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "5T")
        End If
        Exit Sub
    Case "EQUI"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
            If UserList(UserIndex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call EquiparInvItem(UserIndex, val(rdata))
        Exit Sub

    Case "CHEA"
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) > 0 And val(rdata) < 5 Then
            UserList(UserIndex).Char.Heading = rdata
            Call ChangeUserChar(ToPCAreaG, UserIndex, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
        Exit Sub

    Case "SKSE"
        Dim sumatoria As Integer
        Dim incremento As Integer
        rdata = Right$(rdata, Len(rdata) - 4)



        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))

            If incremento < 0 Then

                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                UserList(UserIndex).Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If

            sumatoria = sumatoria + incremento
        Next

        If sumatoria > UserList(UserIndex).Stats.SkillPts Then


            Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If


        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
            UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento
            If UserList(UserIndex).Stats.UserSkills(i) > 100 Then UserList(UserIndex).Stats.UserSkills(i) = 100
        Next
        Exit Sub
    Case "ENTR"

        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub

        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub

        rdata = Right$(rdata, Len(rdata) - 4)

        If Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas < MAXMASCOTASENTRENADOR Then
            If val(rdata) > 0 And val(rdata) < Npclist(UserList(UserIndex).flags.TargetNpc).NroCriaturas + 1 Then
                Dim SpawnedNpc As Integer
                SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNpc).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(UserIndex).flags.TargetNpc).POS, True, False)
                If SpawnedNpc <= MAXNPCS Then
                    Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNpc
                    Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas = Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas + 1

                End If
            End If
        Else
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "3Q" & vbWhite & "°" & "No puedo traer más criaturas, mata las existentes!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        End If

        Exit Sub
    Case "COMP"

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        rdata = Right$(rdata, Len(rdata) - 4)
        If UserList(UserIndex).flags.TargetNpc Then

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = NPCTYPE_TIENDA Then
                Call TiendaVentaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(UserIndex).flags.TargetNpc)
                Exit Sub
            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "3Q" & FONTTYPE_TALK & "°" & "No tengo ningún interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
        Else: Exit Sub
        End If


        Call NPCVentaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(UserIndex).flags.TargetNpc)
        Exit Sub
    Case "RETI"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If UserList(UserIndex).flags.TargetNpc Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
        Else: Exit Sub

        End If
        rdata = Right$(rdata, Len(rdata) - 4)
        Call UserRetiraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))

        Exit Sub

    Case "POVE"
        If Npclist(UserList(UserIndex).flags.TargetNpc).flags.TiendaUser Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).flags.TiendaUser <> UserIndex Then Exit Sub
        Else
            Npclist(UserList(UserIndex).flags.TargetNpc).flags.TiendaUser = UserIndex
        End If

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If UserList(UserIndex).flags.TargetNpc Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_TIENDA Then
                Exit Sub
            End If
        Else: Exit Sub
        End If

        rdata = Right$(rdata, Len(rdata) - 4)

        Call UserPoneVenta(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), val(ReadField(3, rdata, 44)))

        Exit Sub

    Case "SAVE"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        If UserList(UserIndex).flags.TargetNpc Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_TIENDA Then Exit Sub
        Else: Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 4)
        Call UserSacaVenta(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
        Exit Sub

    Case "VEND"

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If UserList(UserIndex).flags.TargetNpc Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = NPCTYPE_TIENDA Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "/N")
                Exit Sub
            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "3Q" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 4)
        Call NPCCompraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
        Exit Sub

    Case "DEPO"
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        If UserList(UserIndex).flags.TargetNpc Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
        Else: Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 4)

        Call UserDepositaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
        Exit Sub



    End Select

    Select Case UCase$(Left$(rdata, 5))
    Case "DEMSG"


        If UserList(UserIndex).flags.TargetObj Then
            rdata = Right$(rdata, Len(rdata) - 5)
            Dim f As String, Titu As String, msg As String, f2 As String

            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rdata, 176)
            msg = ReadField(2, rdata, 176)

            Dim n2 As Integer, loopme As Integer
            If FileExist(f, vbNormal) Then
                Dim Num As Integer
                Num = val(GetVar(f, "INFO", "CantMSG"))
                If Num > MAX_MENSAJES_FORO Then
                    For loopme = 1 To Num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
                    Num = 0
                End If

                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & Num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", Num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", 1)
            End If
            Close #n2
        End If
        Exit Sub
    End Select

    Select Case UCase$(Left$(rdata, 6))
    Case "DESCOD"
        rdata = Right$(rdata, Len(rdata) - 6)
        Call UpdateCodexAndDesc(rdata, UserIndex)
        Exit Sub
    Case "DESPHE"
        rdata = Right$(rdata, Len(rdata) - 6)
        Call DesplazarHechizo(UserIndex, CInt(ReadField(1, rdata, 44)), CByte(ReadField(2, rdata, 44)))
        Exit Sub
    Case "PARACE"
        If UserList(UserIndex).flags.Ofreciente = 0 Then Exit Sub
        If Not UserList(UserList(UserIndex).flags.Ofreciente).flags.UserLogged Then Exit Sub

        If NoPuedeEntrarParty(UserList(UserIndex).flags.Ofreciente, UserIndex) Then Exit Sub

        Dim PartyIndex As Integer
        If UserList(UserList(UserIndex).flags.Ofreciente).flags.Party Then
            PartyIndex = UserList(UserList(UserIndex).flags.Ofreciente).PartyIndex
            If PartyIndex = 0 Then Exit Sub
            Call EntrarAlParty(UserIndex, PartyIndex)
        Else
            Call CrearParty(UserIndex)
        End If
        Exit Sub
    Case "PARREC"
        If UserList(UserIndex).flags.Ofreciente = 0 Then Exit Sub
        If Not UserList(UserList(UserIndex).flags.Ofreciente).flags.UserLogged Then Exit Sub
        Call SendData(ToIndex, UserIndex, 0, "||Rechazaste entrar a party con " & UserList(UserList(UserIndex).flags.Ofreciente).Name & "." & FONTTYPE_PARTY)
        Call SendData(ToIndex, UserList(UserIndex).flags.Ofreciente, 0, "||" & UserList(UserIndex).Name & " rechazo entrar en party con vos." & FONTTYPE_PARTY)
        UserList(UserIndex).flags.Ofreciente = 0
        Exit Sub
    Case "PARECH"
        rdata = ReadField(1, Right$(rdata, Len(rdata) - 6), Asc("("))
        rdata = Left$(rdata, Len(rdata) - 1)
        If UserList(UserIndex).flags.Party Then
            If Party(UserList(UserIndex).PartyIndex).NroMiembros = 2 Then
                For i = 1 To Party(UserList(UserIndex).PartyIndex).NroMiembros
                    Call RomperParty(UserIndex)
                Next
            Else
                Call EcharDelParty(NameIndex(rdata))
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No estás en party." & FONTTYPE_PARTY)
        End If
        Exit Sub

    End Select


    Select Case UCase$(Left$(rdata, 7))
    Case "OFRECER"
        rdata = Right$(rdata, Len(rdata) - 7)
        Arg1 = ReadField(1, rdata, Asc(","))
        Arg2 = ReadField(2, rdata, Asc(","))

        If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
            Exit Sub
        End If
        If Not UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then

            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else

            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            End If

            If val(Arg1) = FLAGORO Then

                If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                    Call SendData(ToIndex, UserIndex, 0, "4R")
                    Exit Sub
                End If
            Else

                If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                    Call SendData(ToIndex, UserIndex, 0, "4R")
                    Exit Sub
                End If
                If ObjData(UserList(UserIndex).Invent.Object(val(Arg1)).OBJIndex).NoSeCae Or ObjData(UserList(UserIndex).Invent.Object(val(Arg1)).OBJIndex).Newbie = 1 Or ObjData(UserList(UserIndex).Invent.Object(val(Arg1)).OBJIndex).Real > 0 Or ObjData(UserList(UserIndex).Invent.Object(val(Arg1)).OBJIndex).Caos > 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes ofrecer este objeto." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If UserList(UserIndex).ComUsu.Objeto Then
                Call SendData(ToIndex, UserIndex, 0, "6T")
                Exit Sub
            End If
            UserList(UserIndex).ComUsu.Objeto = val(Arg1)
            UserList(UserIndex).ComUsu.Cant = val(Arg2)
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else

                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto Then

                    UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                    Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "5R" & UserList(UserIndex).Name)
                End If


                Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
        Exit Sub
    End Select


    Select Case UCase$(Left$(rdata, 8))
    Case "ACEPPEAT"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptPeaceOffer(UserIndex, rdata)
        Exit Sub
    Case "PEACEOFF"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call RecievePeaceOffer(UserIndex, rdata)
        Exit Sub
    Case "PEACEDET"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeaceRequest(UserIndex, rdata)
        Exit Sub
    Case "ENVCOMEN"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeticion(UserIndex, rdata)
        Exit Sub
    Case "ENVPROPP"
        Call SendPeacePropositions(UserIndex)
        Exit Sub
    Case "DECGUERR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareWar(UserIndex, rdata)
        Exit Sub
    Case "DECALIAD"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareAllie(UserIndex, rdata)
        Exit Sub
    Case "NEWWEBSI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SetNewURL(UserIndex, rdata)
        Exit Sub
    Case "ACEPTARI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptClanMember(UserIndex, rdata)
        Exit Sub
    Case "RECHAZAR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DenyRequest(UserIndex, rdata)
        Exit Sub
    Case "ECHARCLA"
        Dim eslider As Integer
        rdata = Right$(rdata, Len(rdata) - 8)
        TIndex = NameIndex(rdata)
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
        Call EcharMember(UserIndex, rdata)
        Exit Sub
    Case "ACTGNEWS"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call UpdateGuildNews(rdata, UserIndex)
        Exit Sub
    Case "1HRINFO<"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendCharInfo(rdata, UserIndex)
        Exit Sub
    End Select

    Select Case UCase$(Left$(rdata, 9))
    Case "SOLICITUD"
        rdata = Right$(rdata, Len(rdata) - 9)
        Call SolicitudIngresoClan(UserIndex, rdata)
        Exit Sub
    End Select

    Select Case UCase$(Left$(rdata, 11))
    Case "CLANDETAILS"
        rdata = Right$(rdata, Len(rdata) - 11)
        Call SendGuildDetails(UserIndex, rdata)
        Exit Sub
    End Select

    Procesado = False
End Sub
