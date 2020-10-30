Attribute VB_Name = "Acciones"
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

Public Cruz As Integer
Public Gema As Integer
Option Explicit
Sub ExtraObjs()

    Cruz = UBound(ObjData) - 1
    ObjData(Cruz).Name = "Cruz del Sacrificio"
    ObjData(Cruz).GrhIndex = 116

    Gema = UBound(ObjData)
    ObjData(Gema).Name = "Piedra filosofal incompleta"
    ObjData(Gema).GrhIndex = 705

End Sub
Sub Accion(UserIndex As Integer, Map As Integer, X As Integer, Y As Integer)
    On Error Resume Next

    If Not InMapBounds(X, Y) Then Exit Sub

    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer

    If MapData(Map, X, Y).NpcIndex Then
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 10 Then
            Call SendData(ToIndex, UserIndex, 0, "DL")
            Exit Sub
        End If

        If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_REVIVIR Then
            If UserList(UserIndex).flags.Muerto Then
                Call RevivirUsuarioNPC(UserIndex)
                Call SendData(ToIndex, UserIndex, 0, "RZ")
            Else
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                Call SendUserHP(UserIndex)
            End If
            Exit Sub

        End If

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_BANQUERO Then
            Call IniciarDeposito(UserIndex)
            Exit Sub
        End If

        If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_TIENDA Then
            If Npclist(MapData(Map, X, Y).NpcIndex).flags.TiendaUser > 0 And Npclist(MapData(Map, X, Y).NpcIndex).flags.TiendaUser <> UserIndex Then
                Call IniciarComercioTienda(UserIndex, MapData(Map, X, Y).NpcIndex)
            Else
                Call IniciarAlquiler(UserIndex)
            End If
            Exit Sub
        End If

        If Npclist(MapData(Map, X, Y).NpcIndex).Comercia Then
            Call IniciarComercioNPC(UserIndex)
            Exit Sub
        End If

        If Npclist(MapData(Map, X, Y).NpcIndex).flags.Apostador Then
            UserList(UserIndex).flags.MesaCasino = Npclist(MapData(Map, X, Y).NpcIndex).flags.Apostador
            Call SendData(ToIndex, UserIndex, 0, "ABRU" & UserList(UserIndex).flags.MesaCasino)
            Exit Sub
        End If

        If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Exit Sub
        End If

        If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_VIEJO Then
            If (UserList(UserIndex).Stats.ELV >= 40 And UserList(UserIndex).Stats.RecompensaLevel <= 2) Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 4 Then
                    Call SendData(ToIndex, UserIndex, 0, "DL")
                    Exit Sub
                End If
            End If
            If Not ClaseBase(UserList(UserIndex).Clase) And Not ClaseTrabajadora(UserList(UserIndex).Clase) And UserList(UserIndex).Clase <= GUERRERO Then
                Call SendData(ToIndex, UserIndex, 0, "RELOM" & UserList(UserIndex).Clase & "," & UserList(UserIndex).Stats.RecompensaLevel)
                Exit Sub
            End If
        End If

        If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_NOBLE Then
            If ClaseBase(UserList(UserIndex).Clase) Or ClaseTrabajadora(UserList(UserIndex).Clase) Then Exit Sub

            If UserList(UserIndex).Faccion.Bando <> Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion Then
                Call SendData(ToIndex, UserIndex, 0, Mensajes(Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion, 16) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If

            If UserList(UserIndex).Faccion.Jerarquia = 0 Then
                Call Enlistar(UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion)
            Else
                Call Recompensado(UserIndex)
            End If

            Exit Sub
        End If
    End If


    If MapData(Map, X, Y).OBJInfo.OBJIndex Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.OBJIndex

        Select Case ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType

        Case OBJTYPE_PUERTAS
            Call AccionParaPuerta(Map, X, Y, UserIndex)
        Case OBJTYPE_CARTELES
            Call AccionParaCartel(Map, X, Y, UserIndex)
        Case OBJTYPE_FOROS
            Call AccionParaForo(Map, X, Y, UserIndex)
        Case OBJTYPE_LEÑA
            If MapData(Map, X, Y).OBJInfo.OBJIndex = FOGATA_APAG Then
                Call AccionParaRamita(Map, X, Y, UserIndex)
            End If
        Case OBJTYPE_ARBOLES
            Call AccionParaArbol(Map, X, Y, UserIndex)

        End Select

    ElseIf MapData(Map, X + 1, Y).OBJInfo.OBJIndex Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.OBJIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).ObjType

        Case 6
            Call AccionParaPuerta(Map, X + 1, Y, UserIndex)

        End Select
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).ObjType

        Case 6
            Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)

        End Select
    ElseIf MapData(Map, X, Y + 1).OBJInfo.OBJIndex Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.OBJIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).ObjType & "," & ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).ObjType

        Case 6
            Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

        End Select

    Else
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
    End If

    If MapData(Map, X, Y).Agua = 1 Then Call AccionParaAgua(Map, X, Y, UserIndex)

End Sub
Sub AccionParaRamita(Map As Integer, X As Integer, Y As Integer, UserIndex As Integer)
    On Error Resume Next
    Dim Suerte As Byte
    Dim exito As Byte
    Dim Obj As Obj
    Dim raise As Integer, nPos As WorldPos

    nPos.Map = Map
    nPos.X = X
    nPos.Y = Y

    If Distancia(nPos, UserList(UserIndex).POS) > 4 Then
        Call SendData(ToIndex, UserIndex, 0, "DL")
        Exit Sub
    End If

    If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
        Suerte = 3
    ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
        Suerte = 2
    ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
        Suerte = 1
    End If

    exito = RandomNumber(1, Suerte)

    If exito = 1 Then
        Obj.OBJIndex = FOGATA
        Obj.Amount = 1

        Call SendData(ToIndex, UserIndex, 0, "7O")
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "FO")

        Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)


    Else
        Call SendData(ToIndex, UserIndex, 0, "8O")
    End If


    If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
        Call SubirSkill(UserIndex, Supervivencia)
    End If

End Sub
Sub AccionParaAgua(Map As Integer, X As Integer, Y As Integer, UserIndex As Integer)

    If MapData(Map, X, Y).Agua = 0 Then Exit Sub

    If UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 75 And UserList(UserIndex).Stats.MinAGU < UserList(UserIndex).Stats.MaxAGU Then
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = Minimo(UserList(UserIndex).Stats.MinAGU + 10, UserList(UserIndex).Stats.MaxAGU)
        UserList(UserIndex).flags.Sed = 0
        Call SubirSkill(UserIndex, Supervivencia, 75)
        Call SendData(ToIndex, UserIndex, 0, "||Has tomado del agua del mar." & FONTTYPE_INFO)
        Call SendData(ToPCArea, UserIndex, 0, "TW46")
        Call EnviarHyS(UserIndex)
    End If

End Sub
Sub AccionParaArbol(Map As Integer, X As Integer, Y As Integer, UserIndex As Integer)

    If MapData(Map, X, Y).OBJInfo.OBJIndex = 0 Then Exit Sub
    If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType <> OBJTYPE_ARBOLES Then Exit Sub

    If UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 85 And UserList(UserIndex).Stats.MinHam < UserList(UserIndex).Stats.MaxHam Then
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinHam = Minimo(UserList(UserIndex).Stats.MinHam + 10, UserList(UserIndex).Stats.MaxHam)
        UserList(UserIndex).flags.Hambre = 0
        Call SubirSkill(UserIndex, Supervivencia, 75)
        Call SendData(ToIndex, UserIndex, 0, "||Has comido de los frutos del árbol." & FONTTYPE_INFO)
        Call SendData(ToPCArea, UserIndex, 0, "TW7")
        Call EnviarHyS(UserIndex)
    End If

End Sub
Sub AccionParaForo(Map As Integer, X As Integer, Y As Integer, UserIndex As Integer)
    On Error Resume Next


    Dim f As String, tit As String, men As String, Base As String, auxcad As String
    f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ForoID) & ".for"
    If FileExist(f, vbNormal) Then
        Dim Num As Integer
        Num = val(GetVar(f, "INFO", "CantMSG"))
        Base = Left$(f, Len(f) - 4)
        Dim i As Integer
        Dim N As Integer
        For i = 1 To Num
            N = FreeFile
            f = Base & i & ".for"
            Open f For Input Shared As #N
            Input #N, tit
            men = ""
            auxcad = ""
            Do While Not EOF(N)
                Input #N, auxcad
                men = men & vbCrLf & auxcad
            Loop
            Close #N
            Call SendData(ToIndex, UserIndex, 0, "FMSG" & tit & Chr$(176) & men)

        Next
    End If
    Call SendData(ToIndex, UserIndex, 0, "MFOR")
End Sub


Sub AccionParaPuerta(Map As Integer, X As Integer, Y As Integer, UserIndex As Integer)
    On Error Resume Next

    Dim MiObj As Obj
    Dim wp As WorldPos

    If Not (Distance(UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y, X, Y) > 2) Then
        If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Llave = 0 Then
            If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Cerrada Then

                If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Llave = 0 Then

                    MapData(Map, X, Y).OBJInfo.OBJIndex = ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).IndexAbierta

                    Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)


                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0


                    Call Bloquear(ToMap, 0, Map, Map, X, Y, 0)
                    Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 0)



                    SendData ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_PUERTA

                Else
                    Call SendData(ToIndex, UserIndex, 0, "9O")
                End If
            Else

                MapData(Map, X, Y).OBJInfo.OBJIndex = ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).IndexCerrada

                Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)


                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1


                Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 1)
                Call Bloquear(ToMap, 0, Map, Map, X, Y, 1)

                SendData ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_PUERTA
            End If

            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.OBJIndex
        Else
            Call SendData(ToIndex, UserIndex, 0, "9O")

        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "DL")
    End If

End Sub
Sub AccionParaCartel(Map As Integer, X As Integer, Y As Integer, UserIndex As Integer)
    On Error Resume Next

    Dim MiObj As Obj

    If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = 8 Then

        If Len(ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Texto) > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "MCAR" & _
                                                 ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Texto & _
                                                 Chr$(176) & ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).GrhSecundario)
        End If

    End If

End Sub

